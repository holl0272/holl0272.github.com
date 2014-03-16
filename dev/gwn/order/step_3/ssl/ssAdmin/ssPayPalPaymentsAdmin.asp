<%Option Explicit
'********************************************************************************
'*   PayPal Payments															*
'*   Release Version:   3.00.002												*
'*   Release Date:		March 17, 2003											*
'*   Revision Date:		November 15, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   3.00.002 (November 15, 2004)					                            *
'*   - Modified SummaryFilter to set intial filter to today's payments only     *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

'***********************************************************************************************

Function SummaryFilter

Dim pstrOrderBy
Dim pstrsqlWhere

	mbytText_Filter = Request.Form("optText_Filter")
	mstrText_Filter = Request.Form("Text_Filter")
	
	'load the radio filters
	mbytStatus_Filter = Request.Form("optStatus_Filter")
	mbytShipment_Filter = Request.Form("optShipment_Filter")
	mbytDate_Filter = Request.Form("optDate_Filter")
	mbytoptCategory_Filter	= Request.Form("optCategory_Filter")
	
	mstrStartDate = Request.Form("StartDate")
	mstrEndDate = Request.Form("EndDate")
	
	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")
	
	'set the defaults
	If len(mbytStatus_Filter) = 0 Then mbytShipment_Filter = 0
	If len(mbytShipment_Filter) = 0 Then mbytShipment_Filter = 1
	If len(mbytDate_Filter) = 0 Then mbytDate_Filter = 0
	If len(mbytoptCategory_Filter) = 0 Then mbytoptCategory_Filter = 2
	If len(mstrOrderBy) = 0 Then mstrOrderBy = 5
	If len(mstrSortOrder) = 0 Then mstrSortOrder = "Desc"
	
	If Len(Request.Form) = 0 Then
		mstrStartDate = DateAdd("d", -1, Date())
	End If

	'Now build the SQL
	pstrsqlWhere = " Where (payer_email<>'')"
	pstrsqlWhere = " Where (not payer_email is Null)"

	'load the text filter
	If len(mstrText_Filter) > 0 Then
		Select Case mbytText_Filter
			Case "0"	'Do Not Include
			Case "1"	'Item Name
				pstrsqlWhere = pstrsqlWhere & " AND (item_name Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "2"	'Last Name
				pstrsqlWhere = pstrsqlWhere & " AND (last_name Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "3"	'Buyer Email
				pstrsqlWhere = pstrsqlWhere & " AND (payer_email Like '%" & mstrText_Filter & "%')"
		End Select	
	End If

	If len(mstrStartDate) > 0 then 
		if cblnSQLDatabase Then
			pstrsqlWhere = pstrsqlWhere & " and (payment_date >= '" & mstrStartDate & " 12:00:00 AM')"
		Else
			pstrsqlWhere = pstrsqlWhere & " and (payment_date >= #" & mstrStartDate & " 12:00:00 AM#)"
		End If
	End If
	
	If len(mstrEndDate) > 0 then 
		If cblnSQLDatabase Then
			pstrsqlWhere = pstrsqlWhere & " and (payment_date <= '" & mstrEndDate & " 11:59:59 PM')"
		Else
			pstrsqlWhere = pstrsqlWhere & " and (payment_date <= #" & mstrEndDate & " 11:59:59 PM#)"
		End If
	End If


	Select Case mbytStatus_Filter
		Case "0"	'Do Not Include
		Case "1"	'Completed
			pstrsqlWhere = pstrsqlWhere & " and payment_status='Completed'"
		Case "2"	'Pending
			pstrsqlWhere = pstrsqlWhere & " and payment_status='Pending'"
		Case "3"	'Failed
			pstrsqlWhere = pstrsqlWhere & " and payment_status='Failed'"
		Case "4"	'Denied
			pstrsqlWhere = pstrsqlWhere & " and payment_status='Denied'"
	End Select	

	Select Case mbytoptCategory_Filter
		Case "0"	'Do Not Include
		Case "1"	'Flagged
			pstrsqlWhere = pstrsqlWhere & " and Category=1"
		Case "2"	'unflagged
			pstrsqlWhere = pstrsqlWhere & " and (Category=0 or Category is Null)"
	End Select	

	'Build  the Order By
	dim paryOrderBy(9)
	paryOrderBy(0) = "Category"	'Default
	paryOrderBy(1) = "item_name"
	paryOrderBy(2) = "item_number"
	paryOrderBy(3) = "quantity"
	paryOrderBy(4) = "last_name"
	paryOrderBy(5) = "payment_date"
	paryOrderBy(6) = "payment_gross"
	paryOrderBy(7) = "payment_fee"
	paryOrderBy(8) = "(payment_gross-payment_fee)"
	paryOrderBy(9) = "payment_status"

	pstrOrderBy = " Order By " & paryOrderBy(mstrOrderBy) & " " & mstrSortOrder 
'debugprint "SummaryFilter",	pstrsqlWhere  & pstrOrderBy
	SummaryFilter = pstrsqlWhere  & pstrOrderBy
	
End Function    'SummaryFilter


'If Len(Session("login")) = 0 Then Response.Redirect "Admin.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME")
mstrPageTitle = "PayPalPayments Administration"

%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssPayPalPayments_class.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="../SFLib/ssPayPal_IPNCustomActions.asp"-->
<!--#include file="../SFLib/ssPayPal_IPNDebugModule.asp"-->
<%
'<!--#include file="SSLibrary/modDatabase.asp"-->
'Assumptions:
'   Connection: defines a previously opened connection to the database

'To convert from one naming convention to another:
Dim pobjConn
Set pobjConn = cnn

'page variables
Dim mAction
Dim mstrTXID
Dim mlngTableID

Dim mblnShowFilter, mblnShowSummary
Dim mstrsqlWhere, mstrSortOrder,mstrOrderBy
'Filter Elements
Dim mbytText_Filter
Dim mstrText_Filter

Dim mstrStartDate, mstrEndDate

Dim mbytShipment_Filter
Dim mbytStatus_Filter
Dim mbytDate_Filter
Dim mbytoptCategory_Filter
Dim mbytoptBackOrdered

	mAction = LoadRequestValue("Action")
	mstrTXID = LoadRequestValue("txn_id")

    Set mclsPayPalPayments = New clsPayPalPayments

Sub WriteExtraScript
%>
<script language="JavaScript" src="SSLibrary/ssFormValidation.js"></script>
<script language="JavaScript" src="SSLibrary/calendar.js"></script>
<script language="JavaScript">

var theDataForm;

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}

function btnFilter_onclick(theButton)
{

  theDataForm.Action.value = "Filter";
  theDataForm.submit();
  return(true);
}

function btnFileSelected_onclick(theButton)
{

  theDataForm.Action.value = "FileSelected";
  theDataForm.submit();
  return(true);
}

function OpenIPNDetail(strTXID)
{
window.open("ssPayPalPaymentsAdmin.asp?Action=Detail&txn_id="+strTXID,"IPNDetail","toolbar=0,location=0,directories=0,status=0,copyhistory=0,scrollbars=1,width=600,height=600");
}

function DoOnload()
{
theDataForm = document.frmData;
}

//-->
</script>
<center>
<%
End Sub	'WriteExtraScript

	mstrsqlWhere = SummaryFilter
    Select Case mAction
		Case "Filter"
			mclsPayPalPayments.LoadAll
'			Response.Write mclsPayPalPayments.OutputMessage
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Response.Write "<H2>" & mstrPageTitle & "</H2>"
			Call ShowForm(True)
			Call ShowFilter
			Response.Write mclsPayPalPayments.OutputSummary
			Call ShowForm(False)
		Case "FileSelected"
'			Response.Write mclsPayPalPayments.OutputMessage
			mclsPayPalPayments.LoadAll
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Call ShowForm(True)
			Call mclsPayPalPayments.FileSelectedItems
			Response.Write "<H2>" & mstrPageTitle & "</H2>"
			Call ShowFilter
			Response.Write mclsPayPalPayments.OutputSummary
			Call ShowForm(False)
        Case "FileItem"
            mclsPayPalPayments.FileItem mstrTXID, True
            mclsPayPalPayments.Load mstrTXID
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Call ShowDetail
        Case "UnFileItem"
            mclsPayPalPayments.FileItem mstrTXID, False
            mclsPayPalPayments.Load mstrTXID
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Call ShowDetail
		Case "Detail"
            mclsPayPalPayments.Load mstrTXID
			Call WriteHeader("DoOnload();",False)
			Call WriteExtraScript()
			Call ShowDetail
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
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Call ShowDetail
		Case Else
			mclsPayPalPayments.LoadAll
'			Response.Write mclsPayPalPayments.OutputMessage
			Call WriteHeader("DoOnload();",True)
			Call WriteExtraScript()
			Response.Write "<H2>" & mstrPageTitle & "</H2>"
			Call ShowForm(True)
			Call ShowFilter
			Response.Write mclsPayPalPayments.OutputSummary
			Call ShowForm(False)
	End Select
%>
</center>
</body>

</HTML>
<%
    Set mclsPayPalPayments = Nothing
    
	Call ReleaseObject(pobjConn)
	Call ReleaseObject(cnn)

    Response.Flush
%>

<%
Sub ShowForm(blnOpen)

If blnOpen Then %>
<form action='ssPayPalPaymentsAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type="hidden" id="Action" name="Action" value="Update">
<input type="hidden" id="OrderBy" name="OrderBy" value="<%= mstrOrderBy %>">
<input type="hidden" id="SortOrder" name="SortOrder" value="<%= mstrSortOrder %>">
<% Else %>
</form>
<%
End If

End Sub 'ShowForm

Sub ShowFilter
%>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
    <td valign="top">
        <input type="radio" value="1" <% if mbytText_Filter="1" then Response.Write "Checked" %> name="optText_Filter" id="optText_Filter1"><label for=optText_Filter1>Item Name</LABEL><br />
        <input type="radio" value="2" <% if mbytText_Filter="2" then Response.Write "Checked" %> name="optText_Filter" id="optText_Filter2"><label for=optText_Filter2>Last Name</label><br />
        <input type="radio" value="3" <% if mbytText_Filter="3" then Response.Write "Checked" %> name="optText_Filter" id="optText_Filter3"><label for=optText_Filter3>Buyer Email</label><br />
        <input type="radio" value="0" <% if (mbytText_Filter="0" or mbytText_Filter="") then Response.Write "Checked" %> name="optText_Filter" id="optText_Filter0"><label for=optText_Filter0>Do Not Include</label>
        <p>containing the text<br />
        <input type="text" id="Text_Filter" name="Text_Filter" size="20" value="<%= Server.HTMLEncode(mstrText_Filter) %>">
	</td>
<script language="javascript">

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
		theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
		break;
	case "2":
		theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("ww",-1,Date())) %>";
		theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
		break;
	case "3":
		theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("m",-1,Date())) %>";
		theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
		break;
	case "4":
		theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("yyyy",-1,Date())) %>";
		theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
		break;
	case "5":
		theDataForm.StartDate.focus();
		break;
}

}

</script>
    <td valign="top">
        <p>Show Only Orders Placed<br />
        <input type="radio" value="1" <% if mbytDate_Filter="1" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter1"><label for=optDate_Filter1>Day</label>&nbsp;
        <input type="radio" value="2" <% if mbytDate_Filter="2" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter2"><label for=optDate_Filter2>Week</label>&nbsp;
        <input type="radio" value="3" <% if mbytDate_Filter="3" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter3"><label for=optDate_Filter3>Month</label>&nbsp;
        <input type="radio" value="4" <% if mbytDate_Filter="4" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter4"><label for=optDate_Filter4>Year</label>&nbsp;
        <input type="radio" value="5" <% if mbytDate_Filter="5" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter5"><label for=optDate_Filter5>Custom</label>&nbsp;
        <input type="radio" value="0" <% if mbytDate_Filter="0" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" id="optDate_Filter0"><label for=optDate_Filter0>All</label>&nbsp;<br />
        
		<label for="StartDate">Start Date:&nbsp;</label><input id=StartDate name=StartDate value="<%= mstrStartDate %>">
		<a href="javascript:doNothing()" title="Select start date"
		onclick="setDateField(document.frmData.StartDate); top.newWin = window.open('calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img src="images/calendar.gif" border=0></a><br />

		<label for="EndDate">&nbsp;&nbsp;End Date:&nbsp;</label><input id=EndDate name=EndDate value="<%= mstrEndDate %>">
		<a href="javascript:doNothing()" title="Select end date"
		onclick="setDateField(document.frmData.EndDate); top.newWin = window.open('calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img src="images/calendar.gif" border=0></a>
	</td>
    <td valign="top">
        <p>Show payment statuses that are:<br />
        <input type="radio" value="1" <% if mbytStatus_Filter="1" then Response.Write "Checked" %> name="optStatus_Filter" id="optStatus_Filter1"><label for=optStatus_Filter1>Complete</label>&nbsp;
        <input type="radio" value="2" <% if mbytStatus_Filter="2" then Response.Write "Checked" %> name="optStatus_Filter" id="optStatus_Filter2"><label for=optStatus_Filter2>Pending</label>&nbsp;
        <input type="radio" value="3" <% if mbytStatus_Filter="3" then Response.Write "Checked" %> name="optStatus_Filter" id="optStatus_Filter3"><label for=optStatus_Filter3>Failed</label>&nbsp;
        <input type="radio" value="4" <% if mbytStatus_Filter="4" then Response.Write "Checked" %> name="optStatus_Filter" id="optStatus_Filter4"><label for=optStatus_Filter4>Denied</label>&nbsp;
        <input type="radio" value="0" <% if (mbytStatus_Filter="0" or mbytStatus_Filter="") then Response.Write "Checked" %> name="optStatus_Filter" id="optStatus_Filter0"><label for=optStatus_Filter0>All</label>

        <p>Show payments that are:<br />
        <input type="radio" value="1" <% if mbytoptCategory_Filter="1" then Response.Write "Checked" %> name="optCategory_Filter" id="optCategory_Filter1"><label for=optCategory_Filter>Filed</label>&nbsp;
        <input type="radio" value="2" <% if mbytoptCategory_Filter="2" then Response.Write "Checked" %> name="optCategory_Filter" id="optCategory_Filter2"><label for=optCategory_Filter>Unfiled</label>&nbsp;
        <input type="radio" value="0" <% if (mbytoptCategory_Filter="0" or mbytoptCategory_Filter="") then Response.Write "Checked" %> name="optCategory_Filter" id="optCategory_Filter0"><label for=optCategory_Filter>Both</label>&nbsp;
		<p>
		  <input class="butn" id=btnFileSelected name=btnFileSelected type=button value="FileSelected" onclick="btnFileSelected_onclick(this);">&nbsp;
		  <input class="butn" id="btnFilter" name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);">
		</p>
	</td>
  </tr>
</table>
<%
End Sub 'ShowFilter

'**************************************************************************************************************************************************************************************

Sub ShowDetail

    With mclsPayPalPayments
%>
<script language=javascript>
<!--

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete this transaction?");
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

function fileTransaction(blnFile)
{

    if (blnFile)
    {
    document.frmData.Action.value = "FileItem";
    }
    else
    {
    document.frmData.Action.value = "UnFileItem";
    }
    document.frmData.submit();
    return(true);

}

//-->
</script>

<body>
<center>
<%= .OutputMessage %>

<form action='ssPayPalPaymentsAdmin.asp' id="Form1" name=frmData onsubmit='return ValidInput();' method=post>
<input type="hidden" id="txn_id" name="txn_id" value="<%= .txn_id %>">
<input type=hidden id="Hidden1" name=Action value='Update'>
<table border=1 cellpadding=3 cellspacing=0 id="Table1">
  <colgroup valign = top align=right />
  <colgroup valign = top align=left />
<%
			Dim pstrTemp
			Select Case Trim(.txn_type & "")
				Case "web_accept"
					pstrTemp = "<b>Web Accept Payment Received</b>"
				Case "cart"
					pstrTemp = "<b>Cart Payment Received</b>"
				Case "send_money"
					pstrTemp = "<b>Sent Money Received</b>"
			End Select
%>
 
      <tr>
        <td colspan="2" align="center"><%= pstrTemp %>&nbsp;(<a href="https://www.paypal.com/vst/id=<%= .txn_id %>" title="View details at PayPal.com" target="_blank"><%= .txn_id %></a>)</td>
      </tr>
      <tr>
        <td>Payment Date:</td>
        <td><%= .payment_date %></td>
      </tr>
      <tr>
        <td>Payment:</td>
        <td>
			<table cellpadding="2" cellspacing="0" border="0" id="Table2">
				<tr><td align="right">Gross:</td><td align="right"><%= WriteCurrency(.payment_gross) %></td></tr>
				<tr><td align="right">Fee:</td><td align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_fee) %></td></tr>
				<tr><td align="right">Net:</td><td align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_gross - .payment_fee) %></td></tr>
			</table>
        </td>
      </tr>
      <% If .Recordset.RecordCount = 1 Then %>
      <tr>
        <td>Payment Status</LABEL></td>
        <td><%= .payment_status %>&nbsp;<% If Len(.pending_reason & "") > 0 Then Response.Write "(" & .pending_reason & ")" %></td>
      </tr>
      <% Else %>
      <tr>
        <td>Payment Status</LABEL></td>
        <td>
      </tr>
      <% 
			Do While Not .Recordset.EOF
				Response.Write .Recordset.Fields("IPNpaymentStatus").Value
				If Len(.Recordset.Fields("IPNpendingReason").Value & "") > 0 Then Response.Write "(" & .Recordset.Fields("IPNpendingReason").Value & ")&nbsp;"
				Response.Write .Recordset.Fields("IPNpendingReason").Value	'PayPalIPNs.DateIPNReceived
				.RecordSet.MoveNext
			Loop
			.RecordSet.MoveFirst
%>
        </TD>
      </TR>
<%
         End If 
      %>
      <tr>
        <td>Payment Type</LABEL></td>
        <td><% 
        If (.payment_type="instant") Then 
			Response.Write "Instant" 
		ElseIf .payment_type="echeck" Then 
			Response.Write "eCheck" 
		End If %></td>
      </tr>
      <tr>
        <td>Paid by:</td>
        <td>
        <%
			Dim pstrpayer_status
			Select Case Trim(.payer_status & "")
				Case "verified"
					pstrpayer_status = "The sender of this payment is Verified"
				Case "unverified"
					pstrpayer_status = "The sender of this payment is <b>Unverified</b>"
				Case "intl_verified"
					pstrpayer_status = "The sender of this payment is <i>International</i> Verified"
				Case "intl_unverified"
					pstrpayer_status = "The sender of this payment is <i>International</i> <b>Unverified</b>"
			End Select
		%>
			<%= .first_name %>&nbsp;<%= .last_name %>&nbsp;(<%= pstrpayer_status %>)<br/>
			<%= .address_street %><br/>
			<%= .address_city %>,&nbsp;<%= .address_state %>&nbsp;<%= .address_zip %><br/>
			<%= .address_country %><br/>
			<% If .address_status = "confirmed" Then %>
			This address is confirmed
			<% Else %>
			This address is unconfirmed
			<% End If %>
			<br/>
			<a href="mailto:<%= .payer_email %>"><%= .payer_email %></a><br/>
        </td>
      </tr>
      <tr>
        <td>Quantity:</td>
        <td><%= .quantity %></td>
      </tr>
      <tr>
        <td>Item/Product Name:</td>
        <td><%= .item_name %></td>
      </tr>
      <tr>
        <td>Item/Product Number:</td>
        <td><%= .item_number %></td>
      </tr>
      <% If Len(.invoice & "") > 0 Then %>
      <tr>
        <td>Invoice:</td>
        <td><%= .invoice %>&nbsp;</td>
      </tr>
      <% End If %>
      <% If Len(.custom & "") > 0 Then %>
      <tr>
        <td>Custom:</td>
        <td><%= .custom %>&nbsp;</td>
      </tr>
     <% End If %>
     <% Call WriteCustomActionTable %>
      <tr>
        <td>&nbsp;</td>
        <td>
        <% If .Category=1 Then %>
        <a href="" onclick="fileTransaction(false); return false;">Unfile this transaction</a>
        <% Else %>
        <a href="" onclick="fileTransaction(true); return false;">File this transaction</a>
        <% End If %>
        </td>
      </tr>
<!--
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
    </TD>
  </TR>
-->
</table>
<p align="center"><a href="" onclick="javascript: window.close(); return false;">Close Window</a></p>
</form>
<% 

    End With
End Sub	'ShowDetail
%>