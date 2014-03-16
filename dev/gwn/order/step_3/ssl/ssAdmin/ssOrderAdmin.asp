<%Option Explicit
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.002												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		April 12, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 2.00.002 (April 12, 2004)									        *
'*	   - Enhancement: support for displaying notes in order summary added		*
'*																				*
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

%>
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/storeAdminSettings.asp"-->
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="../SFLib/ssmodDownload.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_FilterModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrderDetailModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_ImportModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrdersToXML.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<%

On Error Goto 0	'added because of global error suppression in mail.asp
debug.Enabled = False
debug.RecordTime "Begin Page"

'**************************************************
'
'	Start Code Execution
'

mstrPageTitle = "Order Administration"
'Const cstrSpecial = "Fill Order"
Const cstrSpecial = ""

'page variables
Dim mAction
Dim mlngOrderID
Dim mstrOrderPaymentMethod

'Paging Settings
Dim mblnShowFilter, mblnShowSummary
Dim mstrSortOrder, mstrOrderBy
Dim mlngPageCount,mlngAbsolutePage
Dim mlngMaxRecords

'Misc Settings
Dim mstremailFile

'Display setting
Dim mbytDisplay
Const enDisplayType_OrderDetail = 0
Const enDisplayType_OutstandingPayments = 1
Const enDisplayType_OutStandingShipments = 2
Const enDisplayType_ImportTracking = 3
Const enDisplayType_SalesReport = 4
Const enDisplayType_None = 5

Dim maryExportTemplates

	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	mbytDisplay = LoadRequestValue("optDisplay")
	If len(mbytDisplay) = 0 Then
		mbytDisplay = enDisplayType_OrderDetail
	Else
		mbytDisplay = CLng(mbytDisplay)
	End If
	
    Set mclsOrder = New clsOrder
	debug.RecordSplitTime "Initialization Complete"
    With mclsOrder
    
	mAction = LoadRequestValue("Action")
	mlngOrderID = LoadRequestValue("OrderID")

    Select Case mAction
        Case "Update"
			If Len(Request.Form("payOrderID")) > 0 Then
				.UpdatePayments
			ElseIf Len(Request.Form("trackingNumbersFile")) > 0 Then
				.ImportShipments
			ElseIf Len(Request.Form("shipOrderID")) > 0 Then
				.UpdateShipments
			Else
				.Update
			End If
			.CheckOrderChange
			If cblnAutoShowSummaryOnSave Then mblnShowSummary = False
			If cblnAutoShowFilterOnSave Then mblnShowFilter = False
        Case "DeleteSelected"
			Dim i
			Dim paryOrderIDs
			If isAllowedToDeleteOrder Then
				paryOrderIDs = Split(Request.Form("chkssOrderID"),",")
				For i = 0 To UBound(paryOrderIDs)
					mlngOrderID = paryOrderIDs(i)
					.DeleteOrder mlngOrderID
				Next 'i
			End If	'isAllowedToDeleteOrder
			mlngOrderID = ""
        Case "CreateExport"
			Call exportOrders(Request.Form("chkssOrderID"))
			mlngOrderID = ""
        Case "Delete"
			If isAllowedToDeleteOrder Then
				.DeleteOrder mlngOrderID
				mlngOrderID = ""
				If cblnAutoShowSummaryOnSave Then mblnShowSummary = False
				If cblnAutoShowFilterOnSave Then mblnShowFilter = False
			End If	'isAllowedToDeleteOrder
		Case "PerformCustomAction"
			If Len(mstrTXID) > 0 Then
			Select Case Trim(Request.Form("CustomAction"))
				Case "0": Call PerformCustomAction(mstrTXID,"completed","","", "", True)
				Case "1": Call PerformCustomAction(mstrTXID,"pending","","", "", True)
				Case "2": Call PerformCustomAction(mstrTXID,"failed","","", "", True)
				Case "3": Call PerformCustomAction(mstrTXID,"denied","","", "", True)
			End Select
			End If
            mclsPayPalPayments.Load mstrTXID
        Case "SendConfirmationEmail"
			Call .Load(mlngOrderID)
			Call SendConfirmationEmail(mclsOrder.ssOrderID)
        Case "Special"
			'this is a special function requiring custom work
			'originally added to fill software orders
			Call DoSpecial(mlngOrderID)
		Case "ViewOrder"
			mblnShowFilter = False
			mblnShowSummary = False
		Case "Filter"
			If mbytDisplay = enDisplayType_OrderDetail Then
				mblnShowFilter = False
				mblnShowSummary = True
			Else
				mblnShowFilter = False
				mblnShowSummary = False
			End If
		Case "getOrdersXML"
			'debugprint "chkOrderUID", Request.Cookies("chkOrderUID")
			'Call SendOrdersXMLToResponse(LoadRequestValue("chkOrderUID"))
			'Cookies used since a large report could exceed querystring max
			If Len(LoadRequestValue("chkOrderUID")) > 0 Then
				Call SendOrdersXMLToResponse(LoadRequestValue("chkOrderUID"))
			Else
				Call SendOrdersXMLToResponse(Request.Cookies("chkOrderUID"))
			End If
			mbytDisplay = enDisplayType_None
		Case Else
			mblnShowFilter = cblnShowFilterInitially
			mblnShowSummary = Not cblnShowFilterInitially
			
			If False Then
			'Session.Contents.Remove("ssOrderAdmin_UpdateDefaultStatuses")
			If Len(Session("ssOrderAdmin_UpdateDefaultStatuses")) = 0 Then
				cnn.Execute "Update ssOrderManager Set ssOrderStatus=2 Where ssOrderStatus=1 And ssDateOrderShipped<" & wrapSQLValue(DateAdd("d", -1, Date()), True, enDatatype_date) ,,128
				Response.Write "<h4>External order statuses updated from <em>Shipped</em> to <em>Complete</em> for orders shipped before <em>" & DateAdd("d", -1, Date()) & "</em></h4>"
				'Response.Write "SQL: " & "Update ssOrderManager Set ssOrderStatus=2 Where ssOrderStatus=1 And ssDateOrderShipped<" & wrapSQLValue(DateAdd("d", -1, Date()), True, enDatatype_date) & "<br />"
				Session("ssOrderAdmin_UpdateDefaultStatuses") = 1
			End If
			End If
    End Select
    
    'Call LoadExportTemplates(maryExportTemplates)
	debug.RecordSplitTime "Loading templates . . ."
	Call getFileNamesInFolder(ssAdminPath & "exportTemplates/", ".xsl", maryExportTemplates)
	debug.RecordSplitTime "Templates loaded"
    Select Case mbytDisplay
		Case enDisplayType_OrderDetail:	
			debug.RecordSplitTime "Loading order summaries . . ."
			Call .LoadOrderSummaries(SummaryFilter)
			debug.RecordSplitTime "Loading order detail"
			Call .Load(mlngOrderID)
		Case enDisplayType_OutstandingPayments:	
			debug.RecordSplitTime "Loading outstanding payments . . ."
			Call .LoadOutstandingPayments(SummaryFilter)
		Case enDisplayType_OutStandingShipments:	
			debug.RecordSplitTime "Loading outstanding shipments . . ."
			Call .LoadOutstandingShipments(SummaryFilter)
		Case enDisplayType_ImportTracking:	
			debug.RecordSplitTime "Loading outstanding shipments . . ."
			Call .LoadOutstandingShipments(SummaryFilter)
	End Select
	debug.RecordSplitTime "Loading complete"

	If mbytDisplay <> enDisplayType_None Then
	
	Call WriteHeader("body_onload();",True)

%>
<script LANGUAGE="JavaScript">

var theDataForm;
var strDetailTitle = "<% If len(.ssOrderID) > 0 Then Response.Write "Order Number: " & EncodeString(.ssOrderID,False) %>";
var blnIsDirty;

//tipMessage[...]=[title,text]
tipMessage['ssOrderStatus']=["Data Entry Help","This is the order status the customer sees on the order status page.<br />Options for this field are set in the ssOrderAdmin_common file."]
tipMessage['ssInternalOrderStatus']=["Data Entry Help","This is the order status used for internal processing. The customer does NOT see this.<br />Options for this field are set in the ssOrderAdmin_common file."]
tipMessage['ssDatePaymentReceived']=["Data Entry Help","Double click to set date payment received to today's date."]

function setShipmentDefaults(orderID)
{
//shipDate
//shipVia
//shipTrackingNumber
//shipMail. + orderID

var plngCount = document.frmData.shipOrderID.length;
var i;


	if (plngCount==undefined)
	{
		setElementValue(document.frmData.shipDate, document.frmData.shipDate.tag);
		setElementValue(document.frmData.shipVia, 1);
		setElementValue(document.all("shipMail." + orderID), true);
	}else{
		for (i=0; i < plngCount;i++)
		{
			if (document.frmData.shipOrderID[i].value == orderID)
			{
				setElementValue(document.frmData.shipDate[i], document.frmData.shipDate[i].tag);
				setElementValue(document.frmData.shipVia[i], 1);
				setElementValue(document.all("shipMail." + orderID), true);
			}
		}
	}

}

function setShipDate()
{
	theDataForm.ssDateOrderShipped.value='<%= Date() %>';
	//setElementValue(theDataForm.ssOrderStatus, '2');
	//setElementValue(theDataForm.ssInternalOrderStatus, 1);
}

function MakeSummaryDirty(strID)
{
var pstrField = "isDirty" + strID;
var theField = eval("document.frmData." + pstrField);

theField.value = 1;

}

function MakeDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnReset.disabled = false;
	blnIsDirty = true;
}

var mdicProductID = new ActiveXObject("Scripting.Dictionary");;
<%= setCustomDictionary(mstrProductID, ", ", "ProductID", "product") %>

function body_onload()
{
	theDataForm = document.frmData;
	blnIsDirty = false;
	document.all("spanOrderID").innerHTML = strDetailTitle;
	
	ScrollToElem("selectedSummaryItem");
<% 
If mblnShowSummary Then
	Response.Write "DisplaySection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplaySection('Filter');" & vbcrlf
Else
	Response.Write "DisplaySection('Order');" & vbcrlf
	Response.Write "ScrollToElem('spanOrderID');" & vbcrlf
End If
%>

	FillItem("ProductID");
	showLoadingMessage('');
}

function showLoadingMessage(strMessage)
{

var theLoadingMessage = document.all("loadingMessage");

	if (strMessage.length == 0)
	{
		theLoadingMessage.style.display = "none";
	}else{
		//theLoadingMessage.innerText = strMessage;
		theLoadingMessage.style.display = "";
	}
	window.status = strMessage;
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
function ViewOrder(theValue)
{
	theDataForm.optDisplay[0].checked = true;
	theDataForm.OrderID.value = theValue;
	theDataForm.Action.value = "ViewOrder";
	theDataForm.submit();
	return false;
}
function DoSpecial(theValue)
{
	theDataForm.OrderID.value = theValue;
	theDataForm.Action.value = "Special";
	theDataForm.submit();
	return false;
}

function btnDeleteOrder_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete Order " + theForm.OrderID.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

var blnDeleteProduct = false;
function ValidInput(theForm)
{

// uncheck deleted products as required
if (document.all("deleteodrdtID") != null){if (! blnDeleteProduct){checkAll(theForm.deleteodrdtID, false);}}

//  if (!isNumeric(theForm.prodWeight,false,"Please enter a number for the Order weight.")) {return(false);}

	setProductFilter();

    return(true);
}


function setProductFilter()
{
	var theSelect = theDataForm.ProductID;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
}

function btnFilter_onclick(theButton)
{
	setProductFilter();
	
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return(true);
}

function ReplaceEmailText()
{
	var theString = theDataForm.emailBody.value;
	var r;

	r = theString.replace(/<dateShipped>/i, theDataForm.ssDateOrderShipped.value);
	theString = r.replace(/<shipMethod>/i, theDataForm.ssShippedVia.options[theDataForm.ssShippedVia.selectedIndex].text);

	theDataForm.emailBody.value = theString;
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

function downloadOrders(strCustomAction, strExportTemplate, strExportField)
{

	if (! anyChecked(theDataForm.chkssOrderID))
	{
		alert("Please select at least one order to download.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	if (strExportTemplate == null)
	{
		strCustomAction = 'ssOrderAdmin_Export.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
		strExportTemplate = theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value
	}
	
	if (strCustomAction == '')
	{
		strCustomAction = 'ssOrderAdmin_Export.asp';
	}
	
	if (strExportField == null)
	{
		strExportField = '';
	}
	
	theDataForm.action='ssOrderAdmin_Export.asp';
	theDataForm.Action.value = 'downloadOrders' + '|' + strExportTemplate + '|' + strExportField;
	theDataForm.target='docOrders';
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	return false;
}

function printOrders(strCustomExport)
{

	if (strCustomExport == null)
	{
		strCustomExport = 'ssOrderAdmin_Export.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
	}
	
	if (! anyChecked(theDataForm.chkssOrderID))
	{
		alert("Please select at least one order to print.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	theDataForm.action = strCustomExport;
	theDataForm.Action.value = 'printOrders' + '|' + theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value;
	theDataForm.target='printOrders';
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	return false;

}

function viewOrders(strCustomAction, strExportTemplate, strExportField)
{

	if (! anyChecked(theDataForm.chkssOrderID))
	{
		alert("Please select at least one order.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	if (strExportTemplate == null)
	{
		strCustomAction = 'ssOrderAdmin_Export.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
		strExportTemplate = theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value
	}
	
	if (strCustomAction == '')
	{
		strCustomAction = 'ssOrderAdmin_Export.asp';
	}
	
	if (strExportField == null)
	{
		strExportField = '';
	}
	
	theDataForm.action=strCustomAction;
	theDataForm.Action.value = 'viewOrders' + '|' + strExportTemplate + '|' + strExportField;
	theDataForm.target='docOrders';
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	return false;
}

var strSubSection = "Status";
function DisplaySection(strSection)
{
<% If isArray(mclsOrder.CustomOrderValues) Then %>
var arySections = new Array("Status", "BackOrder", "Order", "Filter", "Summary", "Custom", "ExportStatus");
<% Else %>
var arySections = new Array("Status", "BackOrder", "Order", "Filter", "Summary", "ExportStatus");
<% End If %>
if ((strSection == "Status") || (strSection == "BackOrder") || (strSection == "Custom") || (strSection == "ExportStatus"))
{
	switch (strSection)
	{
		case "Custom":
		{
			document.all("tblStatus").style.display = "none";
			document.all("tdStatus").className = "hdrNonSelected";
			document.all("tblBackOrder").style.display = "none";
			document.all("tdBackOrder").className = "hdrNonSelected";
			document.all("tblExportStatus").style.display = "none";
			document.all("tdExportStatus").className = "hdrNonSelected";
			document.all("tblCustom").style.display = "";
			document.all("tdCustom").className = "hdrSelected";
			break;
		}
		case "Status":
		{
			document.all("tblStatus").style.display = "";
			document.all("tdStatus").className = "hdrSelected";
			document.all("tblBackOrder").style.display = "none";
			document.all("tdBackOrder").className = "hdrNonSelected";
			document.all("tblExportStatus").style.display = "none";
			document.all("tdExportStatus").className = "hdrNonSelected";
			if (document.all("tblCustom") != null)
			{
				document.all("tblCustom").style.display = "none";
				document.all("tdCustom").className = "hdrNonSelected";
			}
			break;
		}
		case "BackOrder":
		{
			document.all("tblStatus").style.display = "none";
			document.all("tdStatus").className = "hdrNonSelected";
			document.all("tblBackOrder").style.display = "";
			document.all("tdBackOrder").className = "hdrSelected";
			document.all("tblExportStatus").style.display = "none";
			document.all("tdExportStatus").className = "hdrNonSelected";
			if (document.all("tblCustom") != null)
			{
				document.all("tblCustom").style.display = "none";
				document.all("tdCustom").className = "hdrNonSelected";
			}
			break;
		}
		case "ExportStatus":
		{
			document.all("tblStatus").style.display = "none";
			document.all("tdStatus").className = "hdrNonSelected";
			document.all("tblBackOrder").style.display = "none";
			document.all("tdBackOrder").className = "hdrNonSelected";
			document.all("tblExportStatus").style.display = "";
			document.all("tdExportStatus").className = "hdrSelected";
			if (document.all("tblCustom") != null)
			{
				document.all("tblCustom").style.display = "none";
				document.all("tdCustom").className = "hdrNonSelected";
			}
			break;
		}
		default:
		{

		}
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

function SendConfirmationEmail(theValue)
{
	theDataForm.OrderID.value = theValue;
	theDataForm.Action.value = "SendConfirmationEmail";
	theDataForm.submit();
	return false;
}

function DeleteSelected()
{
	if (anyChecked(theDataForm.chkssOrderID))
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
//-->
</script>

<center>
<%
	Dim mstrDetailTableTitle
    Select Case mbytDisplay
		Case enDisplayType_OrderDetail:	mstrDetailTableTitle = "Order&nbsp;Detail"
		Case enDisplayType_OutstandingPayments:	mstrDetailTableTitle = "Outstanding&nbsp;Payments"
		Case enDisplayType_OutStandingShipments:	mstrDetailTableTitle = "Outstanding&nbsp;Shipments"
		Case enDisplayType_ImportTracking:	mstrDetailTableTitle = "Import&nbsp;Tracking"
	End Select

	Response.Write .OutputMessage
%>

<form action="ssOrderAdmin.asp" id=frmData name=frmData onsubmit="return ValidInput(this);" method=post>
<input type="hidden" id='OrderID' name='OrderID' value=<%= .ssOrderID %>>
<input type="hidden" id=Action name=Action value="Update">
<input type="hidden" id=blnShowSummary name=blnShowSummary value="">
<input type="hidden" id=blnShowFilter name=blnShowFilter value="">
<input type="hidden" id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type="hidden" id=OrderBy name=OrderBy value="<%= mstrOrderBy %>">
<input type="hidden" id=SortOrder name=SortOrder value="<%= mstrSortOrder %>">

<div id="loadingMessage"><h3><font color=red><blink>Loading . . .</blink></font></h3></div>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplaySection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your order filter criteria here.">&nbsp;Order&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<% If (mbytDisplay = enDisplayType_OrderDetail) Then %>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplaySection('Summary');" onmouseover="window.status = this.title" onmouseout="window.status = ''" title="Orders which meet the filter criteria">&nbsp;Order&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<% End If %>
	<th id="tdOrder" class="hdrNonSelected" nowrap onclick="return DisplaySection('Order');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View results">&nbsp;<%= mstrDetailTableTitle %>&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" id=btnHelp name=btnHelp></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
<% 
	Call ShowFilter
	If (mbytDisplay = enDisplayType_OrderDetail) AND (len(mAction) > 0 or cblnAutoShowTable) Then
		If .OrderSummaryCount > 0 Then Call ShowTemplateSelections
		Response.Write .OutputSummary
	End If
%>
<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblOrder">
  <tr class="tblhdr">
	<th align=center><span id="spanOrderID"></span>&nbsp;</th>
  </tr>
  <tr>
    <td>
	<%
		Select Case mbytDisplay
			Case enDisplayType_OrderDetail:	
				debug.RecordSplitTime "Displaying order detail . . ."
				Call ShowOrderDetail(.rsOrders)
				Call ShowFooter
			Case enDisplayType_OutstandingPayments:	
				Call ShowTemplateSelections
				Call .ShowOutstandingPayments
				Call ShowFooter
				%><script>strDetailTitle = "Outstanding Payments";</script><%
			Case enDisplayType_OutStandingShipments:	
				Call ShowTemplateSelections
				Call .ShowOutstandingShipments
				Call ShowFooter
				%><script>strDetailTitle = "Outstanding Shipments";</script><%
			Case enDisplayType_ImportTracking:	
				Call ShowImportTracking
				Call ShowFooter
				%><script>strDetailTitle = "Import Tracking Information";</script><%
		End Select
		debug.RecordSplitTime "Display complete"
	%>
	</td>
  </tr>
</table>

	</td>
  </tr>
</table>
</form>
</center>
</body>
</HTML>
<%
		End If	'mbytDisplay <> enDisplayType_None
    End With

	debug.WriteSplitTimes
	Call ReleaseObject(cnn)
    Response.Flush
%>

<% Sub ShowFooter() %>
<% If mclsOrder.OrderSummaryCount > 0 Then %>
  <table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
	<tr>
	<td>&nbsp;</td>
	<td>
		<input class='butn' id=btnReset name=btnReset type=reset value=Reset>&nbsp;&nbsp;
		<% If mbytDisplay = enDisplayType_OrderDetail And isAllowedToDeleteOrder Then %>
		<input class='butn' id=btnDeleteOrder name=btnDeleteOrder type=button value='Delete Order' onclick='return btnDeleteOrder_onclick(this)'>
		<% End If %>
		<input class="butn" id="btnUpdate" name="btnUpdate" type="submit" value="Save Changes">
	</td>
	</tr>
<% End If	'mclsOrder.OrderSummaryCount > 0 %>
  </TABLE>
<% End Sub 'ShowFooter %>

<%
'Added for Custom Actions
'Comment the following lines if PayPal IPN not installed
'Dim pobjConn
'<!--#include file="ssPayPalPayments_class.asp"-->
'<!--#include file="../SFLib/ssPayPal_IPNCustomActions.asp"-->
%>
<%
'Comment the above lines if PayPal IPN not installed

Sub ShowCustomPayPalActions

If Instr(1,mstrOrderPaymentMethod, "PayPal") = 0 Then Exit Sub

Set mclsPayPalPayments = New clsPayPalPayments
mclsPayPalPayments.LoadByItemNumber mclsOrder.ssOrderID
debugprint "mstrOrderPaymentMethod",mstrOrderPaymentMethod
%>
 <tr class="tblhdr">
    <th width="100%" colspan="2">Custom Actions</th>
  </tr>
  <tr>
    <td>
      <% Call WriteCustomActionTable %></td></tr><%
End Sub	'ShowCustomPayPalActions

Function DoSpecial(lngOrderID)

Set pobjConn = cnn
Call MyAction(lngOrderID)
%><%
End Function	'DoSpecial

Function SendConfirmationEmail(ByRef lngOrderID)

Dim p_objRS
Dim pstrSQL

Dim pstrEmailTo
Dim pstrEmailBody

Dim pstrSubject
Dim pstrEmailPrimary
Dim pstrEmailSecondary

	'pstrEmailBody = ssCustomOrderEmail(cnn, lngOrderID)

	pstrSQL = "SELECT sfAdmin.adminPrimaryEmail, sfAdmin.adminSecondaryEmail, sfAdmin.adminEmailSubject FROM sfAdmin"
	Set p_objRS = GetRS(pstrSQL)
	If Not p_objRS.EOF Then
		pstrEmailPrimary = p_objRS.Fields("adminPrimaryEmail").Value
		pstrEmailSecondary = p_objRS.Fields("adminSecondaryEmail").Value
		pstrSubject = p_objRS.Fields("adminEmailSubject").Value
	End If
		
'	On Error Resume Next

	pstrSQL = "SELECT sfCustomers.custEmail" _
			& " FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId" _
			& " WHERE orderID=" & lngOrderID
	Set p_objRS = GetRS(pstrSQL)
	If Not p_objRS.EOF Then
		pstrEmailTo = p_objRS.Fields("custEmail").Value
		If Len(pstrEmailTo & "") = 0 Then
			SendConfirmationEmail = False
		Else
			'Prepare string for modified mail routine
			' delimited with |
			'0 - sCustEmail
			'1 - sPrimary	- leave blank to use default, set to - not to send
			'2 - sSecondary	- leave blank to use default, set to - not to send
			'3 - sSubject
			'4 - sMessage

			pstrEmailBody = exportOrders(lngOrderID, Server.MapPath("exportTemplates/" & cstrEmailTemplate))
			Call createMail("",pstrEmailTo & "|" & pstrEmailPrimary & "|" & pstrEmailSecondary & "|" & pstrSubject & "|" & pstrEmailBody)

			SendConfirmationEmail = True
		End If
	End If
	
	p_objRS.Close
	Set p_objRS = Nothing

End Function	'SendConfirmationEmail
%>