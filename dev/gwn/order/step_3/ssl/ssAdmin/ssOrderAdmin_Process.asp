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
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_FilterModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrderDetailModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_ImportModule.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrdersToXML.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssOrderAdmin_Process_Rules.asp"-->
<!--#include file="ssOrderAdmin_Process_Functions.asp"-->
<%

On Error Goto 0	'added because of global error suppression in mail.asp

'**************************************************
'
'	Variable Declarations
'

'page variables
Dim i
Dim maryItem
Dim maryOrders
Dim mlngMaxRecords
Dim mlngRuleCounter
Dim mstrAction
Dim RecordCounter
Dim RuleCounter
Dim mstrReferenceMessage
    
'**************************************************
'
'	Start Code Execution
'

	mstrPageTitle = "Order Processing"
	mstrAction = LoadRequestValue("Action")

	Call WriteHeader("body_onload();",True)

%>
<script LANGUAGE="JavaScript">

var theDataForm;
var blnIsDirty;

//tipMessage[...]=[title,text]
tipMessage['ssOrderStatus']=["Data Entry Help","This is the order status the customer sees on the order status page.<br />Options for this field are set in the ssOrderAdmin_common file."]
tipMessage['ssInternalOrderStatus']=["Data Entry Help","This is the order status used for internal processing. The customer does NOT see this.<br />Options for this field are set in the ssOrderAdmin_common file."]
tipMessage['ssDatePaymentReceived']=["Data Entry Help","Double click to set date payment received to today's date."]

function body_onload()
{
	theDataForm = document.frmData;
	blnIsDirty = false;
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

//-->
</script>

<center>
<%
'	Response.Write .OutputMessage
%>

<form action="ssOrderAdmin_Process.asp" id=frmData name=frmData onsubmit="" method=post>
<input type="hidden" id=Action name=Action value="Update Orders to Manually Process">

<div id="loadingMessage"><h3><font color=red><blink>Loading . . .</blink></font></h3></div>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" id=btnHelp name=btnHelp></th>
  </tr>
  <tr>
	<td class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
<% 
    Select Case mstrAction
        Case "Update Orders to Manually Process"
			Call MassUpdateOrders_SetInternalOrderStatusToFlaggedForManual(Request.Form("MarkForManualProcessing"))
		Case Else
    End Select
    
	If mblnProcessOrdersTestMode Then
		Response.Write "<h1>Test Mode</h1>"
		Response.Write "<p>No actions taken. Click <a href=""ssOrderAdmin_Process.asp?TestMode=Off"">here</a> to perform actions.</p>"
	Else
		Response.Write "<h1>Live Mode</h1>"
		Response.Write "<p>All actions taken. Click <a href=""ssOrderAdmin_Process.asp"">here</a> to return to test mode.</p>"
	End If

%>
<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblOrder">
  <tr>
    <td>
	<%
	
	If loadOrdersToProcess(pbytDaysBackToCheckOrders, mstrReferenceMessage) Then
		Call ProcessOrderActions
	End If
	Call cleanOrderProcessingLog(1)

	'Debug section
	Response.Write "<table class=tbl border=1 cellpadding=2 cellspacing=0>"
	Response.Write "<colgroup><col valign=top /><col valign=top /><col valign=top /><col valign=top /><col valign=top /><col valign=top /><col valign=top align=center /><col valign=top /></colgroup>"
	Response.Write "<tr class=tblhdr><th>Order</th><th>Order Item</th><th>Manufacturer</th><th>Vendor</th><th>Rules</th><th>Actions</th><th>Fraud Score</th><th>Message</th></tr>"
	If isArray(maryOrders) Then

		For RecordCounter = 0 To UBound(maryOrders)
			If Len(maryOrders(RecordCounter)(enProcessingItem_Rules)) = 0 Then
				Response.Write "<tr><td>" & maryOrders(RecordCounter)(enProcessingItem_OrderNumber) & "</td><td colspan=3>No Actions</td></tr>"
			Else
				Response.Write "<tr>"
				Response.Write "<td><input type=checkbox name=MarkForManualProcessing value=" & maryOrders(RecordCounter)(enProcessingItem_OrderNumber) & " title=""Flag this order for manual processing""> " & Replace(pstrOrderDetailLink, "{orderID}", maryOrders(RecordCounter)(enProcessingItem_OrderNumber)) & "</td>"
				'Response.Write "<td>" & Replace(pstrOrderDetailLink, "{orderID}", maryOrders(RecordCounter)(enProcessingItem_OrderNumber)) & "</td>"
				Response.Write "<td>" & maryOrders(RecordCounter)(enProcessingItem_OrderItem) & "</td>"
				Response.Write "<td>" & maryOrders(RecordCounter)(enProcessingItem_OrderItemMfg) & "</td>"
				Response.Write "<td>" & maryOrders(RecordCounter)(enProcessingItem_OrderItemVend) & "</td>"

				maryItem = Split(maryOrders(RecordCounter)(enProcessingItem_Rules), cstrDelimiter)
				Response.Write "<td>"
				For i = 0 To UBound(maryItem)
					Response.Write maryProcessingRules(maryItem(i))(enProcessingRule_Title) & "(" & maryItem(i) & ")<br />"
				Next 'i
				Response.Write "</td>"

				maryItem = Split(maryOrders(RecordCounter)(enProcessingItem_Actions), cstrDelimiter)
				Response.Write "<td>"
				For i = 0 To UBound(maryItem)
					Response.Write maryProcessingActions(maryItem(i))(enProcessingAction_Title) & " (" & maryItem(i) & ")<br />"
				Next 'i
				Response.Write "</td>"

				maryItem = Split(maryOrders(RecordCounter)(enProcessingItem_FraudScore), cstrDelimiter)
				Response.Write "<td>"
				For i = 0 To UBound(maryItem)
					Response.Write maryItem(i) & "<br />"
				Next 'i
				Response.Write "</td>"

				maryItem = Split(maryOrders(RecordCounter)(enProcessingItem_Messages), cstrDelimiter)
				Response.Write "<td>"
				For i = 0 To UBound(maryItem)
					Response.Write maryItem(i) & "<br />"
				Next 'i
				Response.Write "&nbsp;</td>"

				Response.Write "</tr>"
			End If
		Next 'RecordCounter
		Response.Write "<tr><td colspan=""8"" align=""left"">" & UBound(maryOrders) + 1 & " item(s) to process.<br /><input type=submit name=btnSubmit value=""Update Orders to Manually Process""></td></tr>"
	Else
		Response.Write "<tr><td colspan=""8"" align=""center"">No Orders Require Processing</td></tr>"
	End If
	Response.Write "</table>"

'**********************************************************

	%>
	</td>
  </tr>
  <tr>
  <td><%= mstrReferenceMessage %><br /><a href="ssOrderProcessingLogAdmin.asp">Processing Action Log</a></td>
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

	Call ReleaseObject(cnn)
    Response.Flush
%>


