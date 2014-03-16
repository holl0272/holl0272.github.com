<%Option Explicit
'********************************************************************************
'*   Promotion Manager for StoreFront 5.0										*
'*   Release Version:	2.00.001 												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		N/A														*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.				*
'********************************************************************************

Response.Buffer = True

Function Load

'	On Error Resume Next

	set mrsDiscounts = server.CreateObject("adodb.recordset")
'		sql = "SELECT ordersDiscounts.ordersDiscountID, ordersDiscounts.PromotionID, Promotions.PromoCode, Promotions.PromoTitle, ordersDiscounts.OrderID, ordersDiscounts.DiscountAmount " _
'			& "FROM ordersDiscounts INNER JOIN Promotions ON ordersDiscounts.PromotionID = Promotions.PromotionID " _
		sql = "SELECT Promotions.PromoCode, Promotions.PromoTitle, sfOrders.orderDate, ordersDiscounts.ordersDiscountID, ordersDiscounts.OrderID, ordersDiscounts.PromotionID, ordersDiscounts.DiscountAmount " _
			& "FROM (ordersDiscounts INNER JOIN Promotions ON ordersDiscounts.PromotionID = Promotions.PromotionID) INNER JOIN sfOrders ON ordersDiscounts.OrderID = sfOrders.orderID " _
			& mstrsqlWhere

	with mrsDiscounts
		.ActiveConnection = cnn
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = sql
		.Open
	end with

	Load = (err.number = 0)

End Function	'Load

'***********************************************************************************************

Sub OutputSummary()

Dim i
Dim aSortHeader(3,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

	.Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
	.Write "<COLGROUP align='center' width='25%'>"
	.Write "<COLGROUP align='left' width='80%'>"
	.Write "<COLGROUP align='left' width='10%'>"
	.Write "  <tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort by Order Number in descending order"
		aSortHeader(2,0) = "Sort by Promotion Code in descending order"
		aSortHeader(3,0) = "Sort by Discount Amount in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort by Order Number in ascending order"
		aSortHeader(2,0) = "Sort by Promotion Code in ascending order"
		aSortHeader(3,0) = "Sort by Discount Amount in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "Order&nbsp;Number"
	aSortHeader(2,1) = "Promotion&nbsp;Code"
	aSortHeader(3,1) = "Discount&nbsp;Amount"

	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
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

	If mrsDiscounts.RecordCount > 0 Then
		mrsDiscounts.MoveFirst
		For i = 1 To mrsDiscounts.RecordCount
			.Write "<TR>"
			.Write "  <TD><A href='' title='View details of Order" & mrsDiscounts("OrderID").value & "' onclick='document.frmsfReport.OrderID.value=" & mrsDiscounts("OrderID").value & "; document.frmsfReport.submit(); return false;'>" & mrsDiscounts("OrderID").value & "</A></TD>" & vbcrlf
			.Write "  <TD><A href='ssPromotionsAdmin.asp?Action=View&PromotionID=" & mrsDiscounts("PromotionID").value & "' title='View details of Promotion " & mrsDiscounts("PromoCode").value & "'>" & mrsDiscounts("PromoCode").value & "&nbsp;" & mrsDiscounts("PromoTitle").value & "</A></TD>" & vbcrlf
			.Write "  <TD>&nbsp;" & formatcurrency(mrsDiscounts("DiscountAmount").value) & "</TD>" & vbcrlf
			.Write "</TR>" & vbCrLf
			mrsDiscounts.MoveNext
		Next
		If mrsDiscounts.RecordCount > 0 Then
			Response.Write "<TR class='tblhdr'><TH align=center colspan=3>" & mrsDiscounts.RecordCount & " Records.<TH></TR>"
		Else
			Response.Write "<TR class='tblhdr'><TH align=center colspan=3>1 Record.</TH></TR>"
		End If
	Else
		Response.Write "<TR><TD align=center colspan=3><h3>There are no Promotions</h3></TD></TR>"
	End If
	.Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

'***********************************************************************************************

Sub DebugPrint(strField,strValue)
	Response.Write "<H3>" & strField & ": " & strValue & "</H3><br />"
End Sub

'***********************************************************************************************

Sub LoadFilter

dim pstrOrderBy

	'Build Filter
	
	strPromoCode = Request.Form("PromoCode")
	strStartDate = Request.Form("StartDate")
	strEndDate = Request.Form("EndDate")

	if len(strPromoCode) > 0 and strPromoCode <> "- All -" then 
		if len(mstrsqlWhere) = 0 Then
			mstrsqlWhere = "Where PromoCode like '%" & strPromoCode & "%'"
		Else
			mstrsqlWhere = mstrsqlWhere & " and PromoCode like '%" & strPromoCode & "%'"
		End If
	End If
	
	If len(strStartDate) > 0 then 
		if len(mstrsqlWhere) = 0 Then
			mstrsqlWhere = "Where orderDate >= " & wrapSQLValue(strStartDate & " 12:00:00 AM", False, enDatatype_date) & ""
		Else
			mstrsqlWhere = mstrsqlWhere & " and orderDate >= " & wrapSQLValue(strStartDate & " 12:00:00 AM", False, enDatatype_date) & ""
		End If
	End If
	
	if len(strEndDate) > 0 then 
		if len(mstrsqlWhere) = 0 Then
			mstrsqlWhere = "Where orderDate <= " & wrapSQLValue(strEndDate & " 11:59:59 PM", False, enDatatype_date) & ""
		Else
			mstrsqlWhere = mstrsqlWhere & " and orderDate <= " & wrapSQLValue(strEndDate & " 11:59:59 PM", False, enDatatype_date) & ""
		End If
	End If

	'Build the order by clause
	
	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")
	blnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Order Number
			pstrOrderBy = "ordersDiscounts.OrderID"
		Case "2"	'Promotion Code
			pstrOrderBy = "PromoCode"
		Case "3"	'Discount Amount
			pstrOrderBy = "DiscountAmount"
	End Select	

	If len(pstrOrderBy) > 0 then mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & mstrSortOrder
	
End Sub    'LoadFilter

%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim rsPromo
dim mrsDiscounts
dim strPromoCode
dim strStartDate
dim strEndDate
dim sql
dim mstrsqlWhere
dim i
dim strTotalDiscount
dim strTotalOrder
dim strDiscountAmount
Dim mstrOrderBy, mstrSortOrder, blnShowSummary
dim mlngAbsolutePage

	mstrPageTitle = "Promotion Reports"

	strPromoCode = Request.QueryString("PromoCode")
	if len(strPromoCode) = 0 then Call LoadFilter
	Call Load
	
	set rsPromo = server.CreateObject("adodb.recordset")
	with rsPromo
		.ActiveConnection = cnn
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = "Select PromoCode from Promotions Order By PromoCode"
		.Open
	end with


	Call WriteHeader("",True)
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/calendar.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

function SortColumn(strColumn,strSortOrder)
{
	document.frmData.OrderBy.value = strColumn;
	document.frmData.SortOrder.value = strSortOrder;
	document.frmData.submit();
	return false;
}

//-->
</SCRIPT>
<CENTER>
<TABLE border=0 cellPadding=5 cellSpacing=1 width="95%">
  <TR>
    <TH><div class="pagetitle "><%= mstrPageTitle %></div></TH>
    <TH>&nbsp;</TH>
    <TH align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a>
	</TH>
  </TR>
</TABLE>

<FORM action="ssPromotionsReport.asp" id=frmData name=frmData method=post>
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type=hidden id=OrderBy name=OrderBy value="<%= mstrOrderBy %>">
<input type=hidden id=SortOrder name=SortOrder value="<%= mstrSortOrder %>">

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <TR>
    <TD>Select a Promtion</TD>
    <TD>Start Date</TD>
    <TD>End Date</TD>
    <TD>&nbsp;</TD>
  </TR>
  <TR>
    <TD>
		<SELECT id=PromoCode name=PromoCode size=2 style="HEIGHT: 38px; WIDTH: 205px">
<% 	if (len(strPromoCode) = 0) or (strPromoCode = "- All -") then
		Response.Write "<option selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option>- All -</Option>" & vbcrlf
	end if
   for i=1 to rsPromo.RecordCount
		if strPromoCode = trim(rsPromo("PromoCode")) then
			Response.Write "<option selected>" & rsPromo("PromoCode") & "</Option>" & vbcrlf
		else
			Response.Write "<option>" & rsPromo("PromoCode") & "</Option>" & vbcrlf
		end if
	rsPromo.MoveNext
   next 'i
%>
		</SELECT>
	</TD>
    <TD>
		<INPUT id=StartDate name=StartDate Value="<%= strStartDate %>" size="20">
		<A HREF="javascript:doNothing()" title="Select start date"
		onClick="setDateField(document.frmData.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<IMG SRC="images/calendar.gif" BORDER=0></A>
    </TD>
    <TD>
		<INPUT id=EndDate name=EndDate Value="<%= strEndDate %>" size="20">
		<A HREF="javascript:doNothing()" title="Select end date"
		onClick="setDateField(document.frmData.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<IMG SRC="images/calendar.gif" BORDER=0></A>
	</TD>
	<td><INPUT class="butn" id=btnFilter name=btnFilter type=submit value="Apply Filter"></td>
  </TR>
</TABLE>
<br />
<% Call OutputSummary %>
</FORM>
   
<p><input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/PromotionManagerII/help_PromotionManagerII.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
&nbsp;&nbsp;<a href="ssPromotionsAdmin.asp">Promotion Administration</a></p>

<!-- action='../sfReports6.asp' method='post'> -->
<form id="frmsfReport" name="frmsfReport" action="ssOrderAdmin.asp" method="get">
	<input type="hidden" name="Action" value="ViewOrder">
	<input type="hidden" name="optDisplay" value="0">
	<input type="hidden" name="optDate_Filter" value="0">
	<input type="hidden" name="optPayment_Filter" value="0">
	<input type="hidden" name="optShipment_Filter" value="0">
	<input type="hidden" name="OrderID" value="OrderID">
	<input type="hidden" name="btnSubmit.x" id="btnSubmit.x" value="0">
</form>

</CENTER>
</BODY>
</HTML>
<%
set rsPromo = nothing
set mrsDiscounts = nothing
set cnn = nothing
Response.Flush
%>