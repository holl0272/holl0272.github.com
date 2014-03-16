<%Option Explicit
'********************************************************************************
'*   Sales Report for StoreFront 6.0		                                    *
'*   Release Version:	2.00.004		
'*   Release Date:		January 1, 2006
'*   Revision Date:		February 17, 2006
'*
'*   Release Notes:
'*
'*   2.00.004 (February 17, 2006)
'*	 - Common Release
'*                                                                              *
'*   2.00.001 (January 1, 2006)
'*	 - Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.
'********************************************************************************

Response.Buffer = True
Server.ScriptTimeout = 600

'Add-on Settings
Const pstrssAddonCode = "SalesReportSF5"
Const pstrssAddonVersion = "2.00.004"

Const cstrOrdersExtra1_Label = ""
Const cstrInsured_Label = ""
Const cstrPackageWeight_Label = ""
Const cstrOrderTrackingExtra1_Label = ""
Const cblnIncludeProductsInDisplay = False

Dim cstrExportTemplateFolder
Dim mlngMaxRecords
Dim mstrOrderBy
Dim mstrSortOrder
Dim mlngAbsolutePage
Dim mlngPageCount

	'mstrPageTitle = CheckForUpdatedVersion(pstrssAddonCode, pstrssAddonVersion)
	mstrPageTitle = "Advanced Sales Report"
	cstrExportTemplateFolder = "exportTemplates\"

%>
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/storeAdminSettings.asp"-->
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_FilterModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrderDetailModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_ImportModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrdersToXML.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<%

'**************************************************
'
'	Start Code Execution
'

debug.Enabled = False	'True	False
debug.printform

'Display setting
Const enDisplayType_OrderDetail = 0
Const enDisplayType_OutstandingPayments = 1
Const enDisplayType_OutStandingShipments = 2
Const enDisplayType_ImportTracking = 3
Const enDisplayType_None = 4

'Paging Settings
Dim mblnShowFilter:		mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
Dim mblnShowSummary:	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
Dim mstrAction
Dim mlngOrderUID
Dim mstrOrderPaymentMethod
Dim maryTemp
Dim pstrOrderIDs
Dim paryOrderIDs
Dim mstrXSLFilePath
Dim maryExportTemplates
Dim i
Dim paryOrderUIDs
Dim marySalesReportTemplates
Dim mstrSalesReportTemplateToUse
Dim mbytDisplay
Dim cstrDefaultSalesReport
Dim cblnAutoSelectAllOrders

Dim mblnGetCCNumbers
Dim mblnIsOrderManager

	mbytDisplay = enDisplayType_OrderDetail
	mblnGetCCNumbers = False
	mblnIsOrderManager = False

	cstrDefaultSalesReport = getAddonConfigurationSetting("SalesReport", "SalesReportDefaultSalesReport")
	cblnAutoSelectAllOrders = getAddonConfigurationSetting("SalesReport", "SalesReportAutoSelectAllOrders")

	mstrSalesReportTemplateToUse = LoadRequestValue("SalesReportTemplateToUse")
	If Len(mstrSalesReportTemplateToUse) = 0 Then mstrSalesReportTemplateToUse = cstrDefaultSalesReport
	
	Call getFileNamesInFolder(ssAdminPath & cstrExportTemplateFolder, ".xsl", maryExportTemplates)
    Call getFileNamesInFolder(ssAdminPath & cstrExportTemplateFolder & "salesReportTemplates/", ".xsl", marySalesReportTemplates)
	If Len(mstrSalesReportTemplateToUse) = 0 Then mstrSalesReportTemplateToUse = marySalesReportTemplates(0)
	
    Set mclsOrder = New clsOrder
    With mclsOrder
    
    maryTemp = Split(LoadRequestValue("Action"), "|")
    pstrOrderIDs = Request.Form("chkOrderUID")
	If Len(LoadRequestValue("Action")) > 0 Then mstrAction = maryTemp(0)
	If UBound(maryTemp) > 0 Then mstrXSLFilePath = maryTemp(1)
	'mstrAction = LoadRequestValue("Action")

	mlngOrderUID = LoadRequestValue("OrderUID")
	'debugprint "mlngOrderUID", mlngOrderUID
	debug.print "Page Action", mstrAction

	'added for CC decryption
	Dim myPrivateKey
	
	myPrivateKey = LoadRequestValue("myPrivateKey")
	If Len(myPrivateKey) > 0 Then Session("encryptionkey") = myPrivateKey
	
    Select Case mstrAction
		Case "Filter"
			mblnShowFilter = False
			mblnShowSummary = True
        Case "getOrdersXML"
			Response.Write OrdersXML(LoadRequestValue("chkOrderUID"))
			mbytDisplay = enDisplayType_None
		Case Else
			'mblnShowFilter = cblnShowFilterInitially
			'mblnShowSummary = Not cblnShowFilterInitially
    End Select
   
	If mbytDisplay = enDisplayType_None Then
		Call ReleaseObject(cnn)
		Response.Flush
		Response.End
	End If

	If .LoadOrderSummaries(SummaryFilter) Then
	
	End If
	'Call .Load(mlngOrderUID)
	
	Call WriteHeader("body_onload();",True)

	If cblnAutoSelectAllOrders Then
		mblnShowFilter = False
		mblnShowSummary = False
	End If
%>
<script LANGUAGE="JavaScript">

var theDataForm;
var blnDeleteProduct = false;
var blnIsDirty;
var cstrSalesReceiptTemplate = "<%= cstrInvoiceTemplate %>";
var cstrPackingSlipTemplate = "<%= cstrPackingSlipTemplate %>";
var strDetailTitle = "<% If len(.ssOrderID) > 0 Then Response.Write "Order Number: " & EncodeString(.ssOrderID,False) %>";
var strSubSection = "Status";

var mdicProductID = new ActiveXObject("Scripting.Dictionary");;
<% Call setCustomDictionary(mstrProductID, ", ", "ProductID", "product") %>

function body_onload()
{
	theDataForm = document.frmData;
	blnIsDirty = false;
	document.all("spanOrderNumber").innerHTML = strDetailTitle;
	FillItem("ProductID");
	
	ScrollToElem("selectedSummaryItem");
<% 
If mblnShowSummary Then
	Response.Write "DisplaySection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplaySection('Filter');" & vbcrlf
Else
	Response.Write "DisplaySection('Order');" & vbcrlf
	Response.Write "ScrollToElem('spanOrderNumber');" & vbcrlf
End If
%>
	LoadInitialReport();
}

function btnFilter_onclick(theButton)
{
	setProductFilter();
	
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return(true);
}

function btnShowOrderSummary_onclick(theForm)
{
	var ptempTarget = document.frmData.target;
	var ptempAction = document.frmData.action;

	document.frmData.target = "NewWindow";
	document.frmData.action = "ssOrderAdmin_OrderSummary.asp";
	document.frmData.submit();

	document.frmData.target = ptempTarget;
	document.frmData.action = ptempAction;

	return false;
}

function ChangeDate(theOption)
{
	switch (theOption.value)
	{
		case "0":
			theDataForm.StartDate.value= "";
			theDataForm.EndDate.value= "";
			break;
		case "1":
			theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("d",-1,Date())) %>";
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "2":
			theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("ww",-1,Date())) %>";
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "3":
			theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("m",-1,Date())) %>";
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "4":
			theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("yyyy",-1,Date())) %>";
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "5":
			//theDataForm.StartDate.focus();
			break;
	}
}

function ChangeDisplay(blnValue)
{
	theDataForm.StartDate.value = "" ;
	theDataForm.EndDate.value = "" ;
	theDataForm.optText_Filter[theDataForm.optText_Filter.length-1].checked = true ;
	theDataForm.optDate_Filter[5].checked = true ;
	if (blnValue)
	{
		theDataForm.optPayment_Filter[0].checked = true ;
		theDataForm.optShipment_Filter[2].checked = true ;
		}else{
		theDataForm.optPayment_Filter[2].checked = true ;
		theDataForm.optShipment_Filter[0].checked = true ;
	}
}

function checkAll_OrderUIDs(blnCheck)
{
	checkAll(theDataForm.chkOrderUID, blnCheck);
	theDataForm.chkCheckAll.checked = blnCheck;
}

function DeleteSelected()
{
	if (anyChecked(theDataForm.chkOrderUID))
	{
		var blnConfirm = confirm("Are you sure you wish to delete the selected order(s)?");
		if (blnConfirm)
		{
			theDataForm.Action.value = 'DeleteSelected';
			theDataForm.submit();
		}else{
		return false;
		}
	}else{
		alert("Please select at least one order.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	return false;
}

function DisplaySection(strSection)
{
var arySections = new Array("Status", "BackOrder", "Order", "Filter", "Summary");

	if ((strSection == "Status") || (strSection == "BackOrder"))
	{
		if (strSection == "Status")
		{
			document.all("tblStatus").style.display = "";
			document.all("tdStatus").className = "hdrSelected";
			document.all("tblBackOrder").style.display = "none";
			document.all("tdBackOrder").className = "hdrNonSelected";
		}else{
			document.all("tblBackOrder").style.display = "";
			document.all("tdBackOrder").className = "hdrSelected";
			document.all("tblStatus").style.display = "none";
			document.all("tdStatus").className = "hdrNonSelected";
		}
		strSubSection = strSection;
	}else{
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
			document.all("tblSummaryFunctions").style.display = "none";
			switch (strSection)
			{
				case "Summary":
				{
					if (document.all("tblSummary") != null){document.all("tblSummaryFunctions").style.display = "";}
					break;
				}
				case "Order":
				{
					//display only if summary section not present
					if (document.all("tblSummary") == null){document.all("tblSummaryFunctions").style.display = "";}
					break;
				}
			}
		}

		if (document.all("tblStatus") != null)
			{
 			if (strSubSection == "Status")
			{
				document.all("tblStatus").style.display = "";
				document.all("tdStatus").className = "hdrSelected";
				document.all("tblBackOrder").style.display = "none";
				document.all("tdBackOrder").className = "hdrNonSelected";
			}else{
				document.all("tblBackOrder").style.display = "";
				document.all("tdBackOrder").className = "hdrSelected";
				document.all("tblStatus").style.display = "none";
				document.all("tdStatus").className = "hdrNonSelected";
			}
		}
	}	
	 
	return(false);
}

function downloadOrders()
{

var Template = templateSelected();

	if (Template == false){return false;}
	
	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to view.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	var prevAction = theDataForm.Action.value;
	theDataForm.Action.value = 'downloadOrders' + '|' + Template;
	theDataForm.target='docOrders';
	theDataForm.submit();
	theDataForm.target='';
	theDataForm.Action.value = prevAction;
	return false;
}

function exportOrders(strCustomExport)
{

var Template = templateSelected();

	if (Template == false){return false;}

	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to export.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	var prevAction = theDataForm.Action.value;
	theDataForm.action = strCustomExport;
	theDataForm.Action.value = 'exportOrders' + '|' + Template;
	theDataForm.target='docExport';
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	theDataForm.Action.value = prevAction;
	return false;
}

function exportOrdersAccounting(strCustomExport)
{

var Template;

	if (strCustomExport == null)
	{
		strCustomExport = 'ssOrderAdmin.asp';
		Template = templateSelected();
		if (Template == false){return false;}
	}
	
	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to export.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	var prevAction = theDataForm.Action.value;
	theDataForm.action = strCustomExport;
	theDataForm.Action.value = 'exportOrders' + '|' + Template;
	theDataForm.target='docExport';
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	theDataForm.Action.value = prevAction;
	return false;
}

function orderSelected()
{
	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to view.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	return true;
}

function printOrders(strTemplate)
{

	var Template = templateSelected();
	
	if (Template == false)
	{
		if (orderSelected())
		{
			submitForm('printOrders' + '|' + Template, "printOrders")
		}
	}
}

function printthisPage(){ 
 	if (window.print)
 	{
		 window.print() ; 
 	} else {
 		var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
 		document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
 		WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box WebBrowser1.outerHTML = ""; 
 	}
}

function setFormAction(strAction)
{
	var strAction;
	
	if (strAction.length = 0)
	{
		strAction = "ssSalesReportAdvanced.asp";
	}else{
		strAction = "ssOrderAdmin.asp?" + strAction;
	}

	return strAction;
}

function setProductFilter()
{

	var theSelect = theDataForm.ProductID;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
}

function showLoadingMessage(strMessage, blnText)
{
	var blnShowText;
	
	if (blnText)
	{
		blnShowText = true;
	}else{
		blnShowText = false;
	}
	
	if (strMessage.length == 0)
	{
		loadingMessage.style.display = "none";
		window.status = '';
	}else{
		if (blnShowText)
		{
			loadingMessage.innerText = strMessage;
			window.status = strMessage;
		}else{
			loadingMessage.innerHTML = strMessage;
			window.status = '';
		}
		if (frmData.showDebugWindow.checked){writeToOutputWindow(strMessage)};
		loadingMessage.style.display = "";
	}
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function submitForm(strAction, strTarget)
{
	
	var prevAction = theDataForm.Action.value;
	theDataForm.Action.value = strAction;
	theDataForm.target = strTarget;
	theDataForm.submit();
	theDataForm.Action.value = prevAction;
	theDataForm.target='';
	return false;
}

function templateSelected()
{
	if (theDataForm.ExportTemplates.selectedIndex == 0)
	{
		alert("Please select a template.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	return theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value;
}

function ValidInput(theForm)
{

	// uncheck deleted products as required
	if (document.all("deleteodrdtID") != null){if (! blnDeleteProduct){checkAll(theForm.deleteodrdtID, false);}}

	//  if (!isNumeric(theForm.prodWeight,false,"Please enter a number for the Order weight.")) {return(false);}

	setProductFilter();

    return(true);
}

function ViewOrder(theValue)
{
	theDataForm.OrderUID.value = theValue;
	theDataForm.Action.value = "ViewOrder";
	theDataForm.submit();
	return false;
}

function viewOrders(strTemplate)
{
var Template;

	if ((strTemplate == undefined) || (strTemplate == ''))
	{
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
		
		Template = theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value;
	}else{
		Template = strTemplate
	}

	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to view.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	var prevAction = theDataForm.Action.value;
	theDataForm.Action.value = 'viewOrders' + '|' + Template;
	theDataForm.target='docOrders';
	theDataForm.action = setFormAction('ssOrderAdmin_Export.asp');
	theDataForm.submit();
	theDataForm.action = setFormAction('');
	theDataForm.Action.value = prevAction;
	theDataForm.target='';
	return false;
}

function viewOrder_Special(strTemplate)
{
var lngUID = "<%= .ssOrderID %>";

	document.frmDetail.target='docOrders';
	document.frmDetail.Action.value = 'viewOrders' + '|' + strTemplate;
	document.frmDetail.chkOrderUID.value=lngUID;
	document.frmDetail.submit();
	return false;
}

function ViewPage(theValue)
{
	theDataForm.AbsolutePage.value = theValue;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}

//function viewSalesReceipt(){viewOrders(cstrSalesReceiptTemplate);}
//function packingSlip(){viewOrders(cstrPackingSlipTemplate);}

function VoidSelected()
{
	if (anyChecked(theDataForm.chkOrderUID))
	{
		var blnConfirm = confirm("Are you sure you wish to void the selected order(s)?");
		if (blnConfirm)
		{
			theDataForm.Action.value = 'VoidSelected';
			theDataForm.submit();
		}else{
		return false;
		}
	}else{
		alert("Please select at least one order.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	return false;
}

//-->
</script>

<script language="JavaScript">

var ChunkCount = 0; // Counts the chunks as they come through...
var XMLHttp;
var XMLData;        // XMLDOM document for storing data
var XMLDoc;         // XMLDOM document for downloading data
var XSLDoc;         // XSL document to use for transformation
var orderUIDsInReport = '';
var isDataLoaded = false;
var mblnDataLoadInProcess = false;
var maryOrderUIDsToGet;
var mlngStartPos = 0;

var cblnDisplayIntermediateResults = true;
var clngNumOrdersToGet_Auto = 200;
var clngNumOrdersToGet = 20;

function getCheckboxValues(checkBoxes, strSeparator)
{

var plngCount = numCheckboxs(checkBoxes);
var i;
var pstrOut = '';

	if (plngCount == 0){return '';}
	if (plngCount > 1)
	{
		for (i=0; i < plngCount;i++)
		{
			if (checkBoxes[i].checked)
			{
				//divDebugging.innerHTML = divDebugging.innerHTML + i + ': ' + checkBoxes[i].value + '<br />';
				if (checkBoxes[i].value != '')
				{
					if (pstrOut.length == 0)
					{
						pstrOut = checkBoxes[i].value;
					}else{
						pstrOut = pstrOut + strSeparator + checkBoxes[i].value;
					}
				}
			}
		}
	}else{
		if (checkBoxes.checked){pstrOut = checkBoxes.value;}
	}
	
	return pstrOut;
}

function DisplayReportData()
{
	showLoadingMessage('<h4>Processing report initiated</h4>');
	try
	{
		showLoadingMessage('<h4>Processing report . . .</h4>');
		divReport.innerHTML = XMLData.transformNode(XSLDoc);
		DataWindow.document.getElementById("divData").innerHTML = divReport.innerHTML;
		setDataToClipboard();
	}
	catch (e)
	{
		divReport.innerHTML = "<p align=center><h3>No orders meet the filter criteria</h3></p>"
	}
	showLoadingMessage('<h4>Processing report completed</h4>');
	closeOutputWindow();
}

function LoadInitialReport()
{
	document.frmData.NumOrdersToGet_Auto.value = clngNumOrdersToGet_Auto;
	document.frmData.NumOrdersToGet.value = clngNumOrdersToGet;

	if (numCheckboxs(theDataForm.chkOrderUID) <= document.frmData.NumOrdersToGet_Auto.value){checkAll_OrderUIDs(true);}
	LoadReportData();
}
	
function LoadReportData()
{

var pstrOrderIDsToGet;
var pstrURL;
var pstrXSL;
var date = new Date();
var s;

	pstrOrderIDsToGet = getCheckboxValues(theDataForm.chkOrderUID, ',');
	maryOrderUIDsToGet = pstrOrderIDsToGet.split(',');
	//divDebugging.innerHTML = divDebugging.innerHTML + '<fieldset>' + pstrOrderIDsToGet + '</fieldset>';
	//alert('pstrOrderIDsToGet: ' + pstrOrderIDsToGet);
	//alert('maryOrderUIDsToGet: ' + maryOrderUIDsToGet.length);
	//maryOrderUIDsToGet = "";
	pstrXSL = getElementValue(frmData.SalesReportTemplateToUse, '');
	divReport.innerHTML = '<h4>Loading report . . . </h4>';

	if (orderUIDsInReport != pstrOrderIDsToGet)
	{
		XMLData = '';
		s = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getYear();

		orderUIDsInReport = pstrOrderIDsToGet;
		showLoadingMessage('<fieldset><legend>Orders to get</legend>Orders: ' + pstrOrderIDsToGet + '</fieldset>');
		//document.cookie = "chkOrderUID=" + pstrOrderIDsToGet + "; expires=" + s;

		pstrURL = 'ssOrderAdmin.asp?Action=getOrdersXML';

		if (orderUIDsInReport.length > 0)
		{
			window.status = "Loading report template";
			showLoadingMessage('<h4>' + window.status + '</h4>');

			XSLDoc = new ActiveXObject("MSXML2.FreeThreadedDOMDocument");
			XSLDoc.async = false;
			
			XSLDoc.load('exportTemplates/salesReportTemplates/' + pstrXSL);
			var myErr = XSLDoc.parseError;
			if (myErr.errorCode != 0)
			{
				// Show the details of the error
				strOut = "<fieldset><legend>Error loading " + XSLDoc.url + "</legend>" + "Error " + myErr.errorCode + ": " + myErr.reason + "<br />" + "Line " + myErr.line + " Pos " + myErr.linepos + "<br />" + "Text: " + myErr.srcText + "<br />" + "</fieldset>";
				showLoadingMessage(strOut);
			}else{
				//alert(XSLDoc.url);
			}
			
			window.status = "Loading report data";
			showLoadingMessage('<h4>' + window.status + '</h4>');
			LoadReportDataInChunks();
			showLoadingMessage('<hr />');
			showLoadingMessage('');
			closeOutputWindow();
		}
	}else{
		XSLPath = "exportTemplates/salesReportTemplates/" + pstrXSL;
		XSLDoc = new ActiveXObject("MSXML2.FreeThreadedDOMDocument");
		XSLDoc.async = false;
		XSLDoc.load(XSLPath);
		var myErr = XSLDoc.parseError;
		if (myErr.errorCode != 0)
		{
			// Show the details of the error
			strOut = "<fieldset><legend>Error loading " + XSLDoc.url + "</legend>" + "Error " + myErr.errorCode + ": " + myErr.reason + "<br />" + "Line " + myErr.line + " Pos " + myErr.linepos + "<br />" + "Text: " + myErr.srcText + "<br />" + "</fieldset>";
			showLoadingMessage(strOut);
		}else{
			DisplayReportData();
			showLoadingMessage('<hr />');
			showLoadingMessage('');
			closeOutputWindow();
		}
	}

}

function LoadReportDataInChunks()
{

var pstrOrderIDsToGet = '';
var plngNumOrders = maryOrderUIDsToGet.length;
var plngNumChunks = plngNumOrders / document.frmData.NumOrdersToGet.value;
var pcurrentChunk = mlngStartPos / document.frmData.NumOrdersToGet.value;

	plngNumChunks = Math.round(plngNumChunks);
	pcurrentChunk = Math.round(pcurrentChunk);

	//make sure we don't get too many orders
	if ((mlngStartPos + Math.round(frmData.NumOrdersToGet.value)) < plngNumOrders)
	{
		plngNumOrders = mlngStartPos + Math.round(frmData.NumOrdersToGet.value);
	}

	for (var i=mlngStartPos; i < plngNumOrders;i++)
	{
		if (pstrOrderIDsToGet.length == 0)
		{
			pstrOrderIDsToGet = maryOrderUIDsToGet[i];
		}else{
			pstrOrderIDsToGet = pstrOrderIDsToGet + "," + maryOrderUIDsToGet[i];
		}
	}
	mlngStartPos = i;
	
	showLoadingMessage('<fieldset><legend>Orders to get in batch ' + (pcurrentChunk+1) + '</legend>Orders: ' + pstrOrderIDsToGet + '</fieldset>');
	//divDebugging.innerHTML = divDebugging.innerHTML + '<fieldset>' + pstrOrderIDsToGet + '</fieldset>';
	//alert('pstrOrderIDsToGet: ' + pstrOrderIDsToGet);
	//alert('maryOrderUIDsToGet: ' + maryOrderUIDsToGet.length);
	//maryOrderUIDsToGet = "";
	if (pstrOrderIDsToGet.length > 0)
	{
		window.status = "Requesting data " + (pcurrentChunk + 1) + "/" + plngNumChunks;
		showLoadingMessage('<h4 id="requestingData' + pcurrentChunk + '">' + window.status + '</h4>');
		showLoadingMessage('<' + 'script' + '>document.getElementById("requestingData' + pcurrentChunk + '").scrollIntoView(true);<' + '/script' + '>');
		//ScrollOutputWindow('requestingData' + pcurrentChunk);
		RetrieveRemoteData('ssOrderAdmin.asp?Action=getOrdersXML', "chkOrderUID=" + pstrOrderIDsToGet, true)
	}
}

function loadSalesReportTemplate(theSelect)
{
	var selectedSalesReport;
	if (theSelect.selectedIndex > 0)
	{
		showLoadingMessage('<h4>Loading report . . .</h4>');
		LoadReportData();
	}else{
		divReport.innerHTML = "<h4>Select a sales report template</h4>";
	}
}

function mergeProducts()
{

var oProductNodeList = XMLDoc.selectNodes("orders/Products/Product");
var oProductsInOrders = XMLData.selectSingleNode("orders/Products");

var i;
var strXPath;
var plngUID;
var plngNumProductsAdded = 0;

	//showLoadingMessage('<h4>Products in current list: ' + oProductsInOrders.childNodes.length + '</h4>');
	//showLoadingMessage('<h4>Products in new batch: ' + oProductNodeList.length + '</h4>');
	for (i=0;i<oProductNodeList.length;i++)
	{
		showLoadingMessage('Merging products ' + (i+1) + '/' + oProductNodeList.length + '<br />');
		plngUID = oProductNodeList[i].attributes.item(0).nodeValue;
		
		strXPath = "orders/Products/Product[@uid='" + plngUID + "']";
		if (XMLData.selectSingleNode("orders/Products/Product[@uid='" + plngUID + "']") == null)
		{
			oProductsInOrders.appendChild(oProductNodeList[i].cloneNode(true));
			plngNumProductsAdded++;
			//showLoadingMessage('<h4>' + ' product added</h4>');
		}
	}
	
	showLoadingMessage(plngNumProductsAdded + ' product(s) in order merged<br />');
}

function mergeOrders()
{

var oNewOrders = XMLDoc.selectNodes("orders/order");
var oExistingOrders = XMLData.selectSingleNode("orders");

var i;
var plngUID;

	for (i=0;i<oNewOrders.length;i++)
	{
		showLoadingMessage('Merging orders ' + (i+1) + '/' + oNewOrders.length + '<br />');
		plngUID = oNewOrders[i].attributes.item(0).nodeValue;
		
		if (XMLData.selectSingleNode("orders/order[@uid='" + plngUID + "']") == null)
		{
			oExistingOrders.appendChild(oNewOrders[i].cloneNode(true));
		}
	}
	showLoadingMessage('Orders merged<br />');
}

function mergeXMLData()
{

	showLoadingMessage('<h4>Merging new data . . .</h4>');
	showLoadingMessage('<h4>Merging Products . . .</h4>');
	mergeProducts();
	showLoadingMessage('<h4>Merging Orders . . .</h4>');
	mergeOrders();
	showLoadingMessage('<h4>Data merged</h4>');

}

function ReadyStateChange_XMLHTTP()
{
	var strOut = '';
	
	ChunkCount ++;
	if (XMLHttp.readyState == 4)
	{
		isDataLoaded = true;

		XMLDoc = new ActiveXObject("MSXML2.DOMDocument");
		XMLDoc.load(XMLHttp.responseXML);
	
		var myErr = XMLDoc.parseError;
		if (myErr.errorCode != 0)
		{
			// Show the details of the error
			strOut = "<fieldset><legend>Error loading " + XMLDoc.url + "</legend>" + "Error " + myErr.errorCode + ": " + myErr.reason + "<br />" + "Line " + myErr.line + " Pos " + myErr.linepos + "<br />" + "Text: " + myErr.srcText + "<br />" + "</fieldset>";
			showLoadingMessage(strOut);
			window.status = myErr.errorCode;
		}else{

			TransformChunk();
			
			if (mlngStartPos < maryOrderUIDsToGet.length)
			{
				if (cblnDisplayIntermediateResults){DisplayReportData();}
				LoadReportDataInChunks();
			}else{
				mlngStartPos = 0;
				showLoadingMessage('<h4>Data load complete</h4>');
				DisplayReportData();
				window.status = '';
				showLoadingMessage('');
				
				//alert('Report Complete!');
			}
		}
	}else{
		for (var i=0; i <= ChunkCount;i++){strOut = strOut + " .";}
		showLoadingMessage('<p>' + strOut + '</p>');
	}
}

function RetrieveRemoteData(strURL, strFormData, blnPostData)
{
var loadAsync = false;

	XMLHttp = new ActiveXObject("Msxml2.XMLHTTP");
	isDataLoaded = false;
	if (blnPostData)
	{
		//XMLHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
		if (loadAsync){XMLHttp.onreadystatechange = ReadyStateChange_XMLHTTP;}
		XMLHttp.open("POST", strURL, loadAsync);
		XMLHttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		XMLHttp.send(strFormData);
	}else{
		XMLHttp.open("GET", strURL, loadAsync);
		XMLHttp.send();
	}
	if (!loadAsync){ReadyStateChange_XMLHTTP();}
}

function showXMLDebugging()
{
	if (frmData.showDebugWindow.checked)
	{
		frmData.xmlSource.style.display = '';
	}else{
		frmData.xmlSource.style.display = 'none';
	}
}

function TransformChunk()
{
	ChunkCount ++;
	if (XMLDoc.readyState == 4)
	{
		// first pass through so set data node
		showLoadingMessage('Data loaded');
		if (typeof(XMLData) != "object")
		{
			XMLData = new ActiveXObject("MSXML2.DOMDocument");
			XMLData = XMLDoc.cloneNode(true);
		}else{
			mergeXMLData();
		}
		//showLoadingMessage('<textarea>' + XMLDoc.xml + '</textarea>');
		
		isDataLoaded = true;
		frmData.xmlSource.value = XMLData.xml;
	}else{
		showLoadingMessage("<h3><font color=red><blink>Loading . . . " + ChunkCount + "</blink></font></h3>");
	}
}

function useCustomXML()
{
	if (XMLDoc.loadXML(frmData.xmlSource.value))
	{
		alert('new data accepted');
	}else{
		var myErr = XMLDoc.parseError;
		if (myErr.errorCode != 0)
		{
			// Show the details of the error
			strOut = "<fieldset><legend>Error loading " + XMLDoc.url + "</legend>" + "Error " + myErr.errorCode + ": " + myErr.reason + "<br />" + "Line " + myErr.line + " Pos " + myErr.linepos + "<br />" + "Text: " + myErr.srcText + "<br />" + "</fieldset>";
			showLoadingMessage(strOut);
		}else{
			showLoadingMessage('Data retrieved. Processing data . . .');
			TransformChunk();
			DisplayReportData();
			showLoadingMessage('');
		}
	
	}
}
	
var mstrSortOrder = '';
var mstrSortColumn = '';
var mstrSortImage;
var mblnLoop = false;

function onSort(sortColumn, sortOrder, dataType, optImage)
{

	if (mstrSortOrder.length == 0){mstrSortOrder = sortOrder;}
	if (mstrSortColumn.length == 0){mstrSortColumn = sortColumn;}
	
	if (mstrSortImage != null)
	{
		document.getElementById(mstrSortImage).src = "../images/clear.gif"
	//alert(theImage.src);
	}
	
	if (mstrSortColumn == sortColumn)
	{
		if (mstrSortOrder == 'ascending')
		{
			mstrSortOrder = 'descending';
		}else{
			mstrSortOrder = 'ascending';
		}
	}

	if (optImage != null)
	{
		mstrSortImage = optImage;
		if (mstrSortOrder == 'ascending')
		{
			document.getElementById(optImage).src = "images/down.gif"
		}else{
			document.getElementById(optImage).src = "images/up.gif"
		}
	}

	mstrSortColumn = sortColumn;


	try
	{
		XSLDoc.selectSingleNode("//xsl:sort/@select").text = sortColumn;
		XSLDoc.selectSingleNode("//xsl:sort/@order").text = mstrSortOrder;
		XSLDoc.selectSingleNode("//xsl:sort/@data-type").text = dataType;
    }
    catch(e)
    {
		if (mblnLoop)
		{
			//reset
			mblnLoop = false;
		}else{
			mblnLoop = true;
			LoadReportData();
		}
    }
	finally
	{
		DisplayReportData();
	}

}

var DataWindow = openGenericWindow("", "DataWindow");
if (DataWindow.document.getElementById("divData") == null)
{
	writeToGenericWindow('<html><head><link rel="stylesheet" href="ssLibrary/ssStyleSheet.css" type="text/css"></head>', DataWindow);
	writeToGenericWindow('<body>', DataWindow);
	writeToGenericWindow('<div id="divData">Loading Data . . .</div>', DataWindow);
	writeToGenericWindow('</body>', DataWindow);
	writeToGenericWindow('</html>', DataWindow);
}else{
	DataWindow.document.getElementById("divData").innerHTML = "Loading Data . . .";
}

function setDataToClipboard()
{
return false;

	if (DataWindow.document.getElementById("divData") != null)
	{
		DataWindow.document.getElementById("divData").select();
		document.execCommand("Copy")
		
		document.getElementById("divDebugging").select();
		document.execCommand("Copy")
	}
}
</script>

<center>

<%= .writeErrorMessages %>

<form action="" id=frmDetail name=frmDetail method=post>
	<input type="hidden" name="Action" value="">
	<input type="hidden" name="chkOrderUID" value="">
</form>
<form action="ssSalesReportAdvanced.asp" id=frmData name=frmData onsubmit="return ValidInput(this);" method=post>
<input type="hidden" id="OrderUID" name="OrderUID" value="<%= .ssOrderID %>" />
<input type="hidden" id=Action name=Action value="Update" />
<input type="hidden" id=blnShowSummary name=blnShowSummary value="" />
<input type="hidden" id=blnShowFilter name=blnShowFilter value="" />
<input type="hidden" id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>" />
<input type="hidden" id=OrderBy name=OrderBy value="<%= mstrOrderBy %>" />
<input type="hidden" id=SortOrder name=SortOrder value="<%= mstrSortOrder %>" />
<input type="hidden" name="todaysDate" id="todaysDate" value="<%= FormatDateTime(Date()) %>" />

<div id="loadingMessage"><h3><font color=red><blink>Loading . . .</blink></font></h3></div>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplaySection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your order filter criteria here.">&nbsp;Order&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplaySection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Orders which meet the filter criteria">&nbsp;Order&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdOrder" class="hdrNonSelected" nowrap onclick="return DisplaySection('Order');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View results">&nbsp;Report&nbsp;</th>
	<th width="90%" align=right>
	  <span class="pagetitle2"><%= mstrPageTitle %></span>
	  &nbsp;<img src="images/properties.gif" onclick="openProperties('SalesReport')" title="Configure Sales Report Options">
	  &nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/SalesReports/help_SalesReports.htm')" id=btnHelp name=btnHelp title="Release Version <%= pstrssAddonVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
<% 
	Call ShowFilter
	If (len(mstrAction) > 0 or cblnAutoShowTable) Then
		If mclsOrder.OrderSummaryCount > 0 Then Call ShowTemplateSelections
		Response.Write .ShowOrderSummaries(cblnAutoSelectAllOrders)
	End If
%>
<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblOrder">
  <colgroup align="left" width="25%">
  <colgroup align="left" width="75%">
  <tr class="tblhdr">
	<th align=center><span id="spanOrderNumber"></span>&nbsp;</th>
  </tr>
  <tr>
    <td align="right">
		<span id="divSettings" style="display:none">
		<label for="NumOrdersToGet_Auto">Qty to Auto Load:</label>&nbsp;<input name="NumOrdersToGet_Auto" id="NumOrdersToGet_Auto" value="" size="3" title="All orders will be retieved automatically if the number of orders in the summary section is below">&nbsp;
		<label for="NumOrdersToGet">Retrieval Qty:</label>&nbsp;<input name="NumOrdersToGet" id="NumOrdersToGet" value="" size="3" title="Number of orders to retrieve at one time">&nbsp;<input type="checkbox" name="showDebugWindow" id="showDebugWindow" onclick="showXMLDebugging();" title="Display intermediate steps in debugging window" value="1" checked />
		</span>
        <a href="Filter" onclick="showHideElement(divSettings); return false;"><img src="images/filter.bmp" border="0"></a>&nbsp;
        <a href="Refresh" onclick="LoadReportData();return false;"><img src="images/refresh.gif" border="0"></a>&nbsp;
		<select name="SalesReportTemplateToUse" ID="SalesReportTemplateToUse" onchange="loadSalesReportTemplate(this);">
		  <option value="" selected>Select a Template</option>
		<% For i = 0 To UBound(marySalesReportTemplates) %>
		  <option value="<%= marySalesReportTemplates(i) %>" <%= isSelected(marySalesReportTemplates(i)= mstrSalesReportTemplateToUse) %>><%= Replace(marySalesReportTemplates(i), ".xsl", "") %></option>
		<% Next 'i %>
		</select>
    </td>
  </tr>
  <tr>
    <td>
    <div id="divReport">Select a report to view</div>

	<hr>
	<textarea id=xmlSource rows=80 cols=140 style="display: none" onchange="useCustomXML();"></textarea>
	<div id="divDebugging"></div>
	</td>
  </tr>
</table>

	</td>
  </tr>
</table>
</form>
</center>
<!--#include file="AdminFooter.asp"-->
</body>
</HTML>
<%
    End With

	Call ReleaseObject(cnn)
    Response.Flush
%>
