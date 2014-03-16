<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Server.ScriptTimeout = 90		'in seconds. Adjust for large databases. Default is usually 90 seconds

	'Const cstrWebSite = "http://www.yourwebsite.com/OrderHistory.asp"	'URL on site to OrderHistory.asp
	'Const cstrWebSite = "http://www.sandshot.net/OrderHistory.asp"	'URL on site to OrderHistory.asp
	Const cstrWebSite = "../../../OrderHistory.asp"	'URL on site to OrderHistory.asp

	'Order Summary Settings
	Const cblnAutoShowTable = True				'Set to true to automatically load the database summary
	Const cblnShowFilterInitially = False		'Set True to display the filter on the initial page load; False to show the order summary
	Const clngScrollableTableHeight = 400		'Set to height of in-window scrollable window. Set to -1 to display entire summary results using browser scroll bar
	Const clngDefaultRecords = 0				'Set your default Maximum Records to show in summary table; 0 to show all
	Const cblnShowPageNumbers = True			'Set True to display page results by page count; False to display by record count
	Const cblnUseOrderFlags = True				'Set True to display order flags in summary table; False not to
	Const cblnUseBackOrder = True				'Set True to display backorder information; False not to
	Const cblnAutoShowSummaryOnSave = False		'Set True to display the Order Sumary tab after saving the order; False to remain on the Order Detail tab
	Const cblnAutoShowFilterOnSave = False		'Set True to display the Order Filter tab after saving the order; False to remain on the Order Detail tab
	'Const cstrDisplayMemoField = "ssInternalNotes"				'Set True to display the Order Filter tab after saving the order; False to remain on the Order Detail tab
	Const cstrDisplayMemoField = ""				'Set True to display the Order Filter tab after saving the order; False to remain on the Order Detail tab
	Const cstrShippingExportFile = "UPS WorldShip Export.xsl"			'Set to name of template you wish to use for single click access to the export
	Const cstrPaymentExportFile = "AuthorizeNet CC Payments CSV.xsl"			'Set to name of template you wish to use for single click access to the export

	'Order Detail Settings
	Const clngProductNameLength = 0				'Set your default max length of the product name to show on a single line; 0 to show the entire name on one line. Used to prevent side to side scrolling
	Const mblnShowFullCountryName = True		'Set to False to show abbreviation
	cblnShowBackOrderColumn = False				'Set True to always display backorder column; False to display only if present
	cblnShowGiftWrapRow = False					'Set True to always display gift wrap row; False to display only if present
	cblnShowBilledAmountRow = False				'Set True to always display billing rows; False to display only if present
	Const cblnShowInventoryOnHand = True		'Set always display backorder column; False to display only if present
	Const cblnUpdateInventoryOnChanges = True	'Set True to display the filter on the initial page load; False to show the order summary

	'Standard Templates to use for Order Detail links
	Const cstrEmailTemplate_Shipment = "OrderShipped.txt"
	Const cstrEmailTemplate_Payment = "OrderShipped.txt"

	'Standard Templates to use for Order Detail links
	Const cblnUseASPPages = False
	Const cstrEmailTemplate = ""
	Const cstrInvoiceTemplate = "Sales Receipt.xsl"
	Const cstrPackingSlipTemplate = "Packing Slip.xsl"
	
	'Order Status Messages
	'This section must match the section in ssOrderManager
	Dim maryOrderStatuses(7)
	maryOrderStatuses(0) = "Order Placed, Awaiting Payment"
	maryOrderStatuses(0) = ""	'this is set to empty so as not to display anything unless explicity set
	maryOrderStatuses(1) = "Payment Received, Awaiting Shipment"
	maryOrderStatuses(2) = "Payment Received, Will Ship when payment clears"
	maryOrderStatuses(3) = "Awaiting Payment, Awaiting Shipment"
	maryOrderStatuses(4) = "Awaiting Payment, Order Shipped"
	maryOrderStatuses(5) = "Order Shipped"
	maryOrderStatuses(6) = "Order Complete"
	maryOrderStatuses(7) = "Order Cancelled"
	
	'Process Status Messages
	'Display Text, TR class (defined in ssStyleSheets.css), image path (displays next to checkbox)
	Dim maryInternalOrderStatuses(6)
	maryInternalOrderStatuses(0) = Array("Unread", "orderStatus_Unread", "", 0)
	maryInternalOrderStatuses(1) = Array("Read", "orderStatus_Read", "", 1)
	maryInternalOrderStatuses(2) = Array("Pending review", "orderStatus_PendingReview", "", 3)
	maryInternalOrderStatuses(3) = Array("Ready to Ship", "orderStatus_Read", "", 6)
	maryInternalOrderStatuses(4) = Array("Order shipped", "orderStatus_Ordered", "", 2)
	maryInternalOrderStatuses(5) = Array("Fraud", "orderStatus_Fraud", "", 4)
	maryInternalOrderStatuses(6) = Array("Voided", "orderStatus_Void", "", 5)
	
	'Export Settings
	Const cblnAutoExport = True					'Set True to automatically mark orders as exported to the accounting program
	Const cblnUseCustomInvoiceNumber = False	'Set True to use custom invoice numbers (increment for each export); False to use order IDs as invoice numbers
	Const cstrInvoiceOrderPrefix = "WEB-"		'Text to precede invoice numbers
	Const cbytTaxFraction = 8					'Text to use for out of state tax names
	Const cstrTaxEntity_NoTax = "Out of State"	'Text to use for out of state tax names
	Const cstrTaxEntity = "{taxRate} {stateAbbr} State Tax"		'Format to use for tax output:
																'Available Replacements
																' {taxRate}	- taxDecimal in a percent format. Ex. .05 becomes 5%
																' {taxDecimal} - decimal equivalent of the taxRate. Ex. 5% sales tax is .05
																' {stateAbbr} - state abbreviation
																' {countyName} - county name from TaxRate Manager installation
	
	'Compatibility Settings
	cExport_CursorLocation = adUseClient		'Set adUseClient (preferred) or adUseServer (if export results come up blank)
	Const cblnUseSF505Dll = False				'Set True to use 50.5 specific dll, False to use older version
	
	Const ssDebug_Download = False				'Placed here for integration
'/
'/////////////////////////////////////////////////

Dim cblnShowBackOrderColumn
Dim cblnShowGiftWrapRow
Dim cblnShowBilledAmountRow
Dim cExport_CursorLocation

Const enStore_ID = 0
Const enStore_Name = 1
Const enStore_OrderHistoryURL = 2
Const enStore_ReportsURL = 3
Const enStore_PackingSlipURL = 4
Const enStore_EmailDirectory = 5
Const enStore_EmailFrom = 6

Dim maryStores()
Const cstrStoreIDFieldName = ""
Call LoadMultipleStoreSettings

'***********************************************************************************************

	Sub LoadMultipleStoreSettings()

	'/////////////////////////////////////////////////
	'/
	'/  Multiple Store Parameters
	'/

		Exit Sub
		
		ReDim maryStores(1)

'		maryStores(0) = Array("storeID", _
'							  "storeName", _
'							  "URL to OrderHistory.asp, ex. http://www.yourSite.com/OrderHistory.asp", _
'							  "URL to invoice page, ex. http://www.yourSite.com/ssl/ssAdmin/ssOrderAdmin_PrintableDetail.asp", _
'							  "URL to packing slip page, ex. http://www.yourSite.com/ssl/ssAdmin/ssOrderAdmin_PackingSlip.asp", _
'							  "directory off the ssAdmin directory the store specific emails tempates are in, ex. emailMEUTemplates/", _
'							  "email address the email should be from, if left blank the default email will be used")

		maryStores(0) = Array("meu", _
							  "marine-electronics-unlimited", _
							  "http://www.marine-electronics-unlimited.com/OrderHistory.asp", _
							  "http://www.starmarinedepot.com/ssl/admin/sfreports1.asp", _
							  "http://www.starmarinedepot.com/ssl/admin/sfreports1.asp", _
							  "emailMEUTemplates/", _
							  "orders@marine-electronics-unlimited.com")
							  
		maryStores(1) = Array("smd", _
							  "starmarinedepot", _
							  "http://www.starmarinedepot.com/OrderHistory.asp", _
							  "http://www.starmarinedepot.com/ssl/admin/sfreports1.asp", _
							  "http://www.starmarinedepot.com/ssl/admin/sfreports1.asp", _
							  "emailTemplates/", _
							  "orders@starmarinedepot.com")

	'/
	'/////////////////////////////////////////////////

	End Sub	'LoadMultipleStoreSettings

'***********************************************************************************************

	Function ShipMethodsAsOptions(strValue)

	Dim i
	Dim pstrTemp

		pstrTemp = "<option value=''></option>"
		For i=0 to UBound(maryShipMethods)
			If strValue = i Then
				pstrTemp = pstrTemp & "<option value='" & i & "' selected>" & maryShipMethods(i) & "</option>"
			Else
				pstrTemp = pstrTemp & "<option value='" & i & "'>" & maryShipMethods(i) & "</option>"
			End If
		Next
		
		ShipMethodsAsOptions = pstrTemp

	End Function	'ShipMethodsAsOptions

'***********************************************************************************************

%>
<!--#include file="ssShippingMethods_common.asp"-->
