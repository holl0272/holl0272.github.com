<% Option Explicit 
'********************************************************************************
'*   Sales Report							                                    *
'*   Release Version:   1.00.0001												*
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

Response.Buffer = true
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssReports.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<%

	mstrPageTitle = "Sandshot Software WebStore Manager"
	Call CheckLoginStatus_AdminPage(False)
%><!--#include file="adminFooter.asp"--><%
	If Response.Buffer Then Response.Flush
    Call ReleaseObject(cnn)

'*******************************************************************************************************************************************

Function rowStyle(byRef blnEvenRow)
	If blnEvenRow Then
		rowStyle = "class='Inactive'"
	Else
		rowStyle = ""
	End If
	blnEvenRow = Not blnEvenRow
End Function	'rowStyle

'*******************************************************************************************************************************************

Sub ShowAdminPage

Dim pstrOrderDetailLink
Dim pdblRunningOrderTotal
Dim i
Dim maryWorking
Dim pblnEvenRow

'Link to sfReports
pstrOrderDetailLink = "<a href='../sfReports1.asp?OrderID=<orderID>'><orderID></a>"
pblnEvenRow = False
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center" ID="Table2">
  <tr>
    <td width="100%" valign="top" align="center">
	<!-- Begin Content Section -->
	<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
      <tr>
        <td align="left" valign="top">
	<!--webbot bot="PurpleText" preview="Begin Content Section" -->
	
<!--webbot bot="PurpleText" preview="Begin Today's Order" -->
<%
If cblnAddon_OrderMgr Then pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&StartDate=" & DateAdd("d", -0, Date()) & "&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'><orderID></a>"
Call getOrderDetails(maryWorking, DateAdd("d", -0, Date()), "", enCriteria_DoNotInclude, enCriteria_DoNotInclude, enCriteria_True)
%>
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblTodaysOrders">
      <tr>
        <th colspan="4" class="hdrNonSelected">Today's Orders</th>
      </tr>
      <% If isArray(maryWorking) Then %>
      <tr class="tdHighlight">
        <th>Order</th>
        <th>Amount</th>
        <th>Customer</th>
        <th>Tx Type</th>
      </tr>
      <%
      pdblRunningOrderTotal = 0
      pblnEvenRow = False
      For i = 0 To UBound(maryWorking)
		If isNumeric(maryWorking(i)(enOrderDetail_orderGrandTotal)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + maryWorking(i)(enOrderDetail_orderGrandTotal)
      %>
      <tr <%= rowStyle(pblnEvenRow) %>>
        <td align=center><%= Replace(pstrOrderDetailLink, "<orderID>", maryWorking(i)(enOrderDetail_orderID)) %>&nbsp;</td>
        <td align=right><%= CustomCurrency(maryWorking(i)(enOrderDetail_orderGrandTotal)) %>&nbsp;</td>
        <td><%= maryWorking(i)(enOrderDetail_custName) %>&nbsp;</td>
        <td><%= maryWorking(i)(enOrderDetail_orderPaymentMethod) %>&nbsp;</td>
      </tr>
      <% Next 'i %>
      <tr>
        <th><%= UBound(maryWorking)+1 %>&nbsp;</th>
        <th align=right><%= CustomCurrency(pdblRunningOrderTotal) %>&nbsp;</th>
        <th colspan=2>&nbsp;</th>
      </tr>
      <% Else %>
      <tr>
        <th colspan="4">There are no orders today</th>
      </tr>
      <% End If	'isArray(maryWorking) %>
    </table>
        <p>
    <!--webbot bot="PurpleText" preview="End Today's Order" --></td>
        <td align="right" valign="top">
<!--webbot bot="PurpleText" preview="Begin Order Summaries" -->
<%
Dim maryDailyOrderSummary
Dim maryWeeklyOrderSummary
Dim maryMonthlyOrderSummary
Dim maryYearlyOrderSummary

Dim mstrYear

Dim mdtToday
Dim mdtFirstDay
Dim mdtFirstDayOfWeek
Dim mdtFirstDayOfMonth
Dim mdtFirstDayOfYear
Dim pstrProductLink

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Const cbytTopSellers = 20
Const cstrOpenTag = "<ol style='MARGIN-LEFT: 18pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;'>"
Const cstrCloseTag = "</ol>"
	
'/
'/////////////////////////////////////////////////

mdtToday = Date()
If True Then
	mstrYear = CStr(Year(mdtToday))
	mdtFirstDay = mdtToday
	mdtFirstDayOfWeek = DateAdd("d", -1 * Weekday(mdtToday, vbSunday) + 1, mdtToday)	'change to vbMonday to start on Sunday
	mdtFirstDayOfMonth = CDate(mstrYear + "/" + CStr(Month(mdtToday)) + "/1")
	mdtFirstDayOfYear = CDate(mstrYear + "/1/1")
Else
	mdtFirstDay = DateAdd("d", -1, mdtToday)
	mdtFirstDayOfWeek = DateAdd("d", -1, mdtToday)
	mdtFirstDayOfMonth = DateAdd("m", -1, mdtToday)
	mdtFirstDayOfYear = DateAdd("y", -1, mdtToday)
End If

maryDailyOrderSummary = GetOrderSummaries(mdtFirstDay, mdtToday, False, cbytTopSellers)
maryWeeklyOrderSummary = GetOrderSummaries(mdtFirstDayOfWeek, mdtToday, False, cbytTopSellers)
maryMonthlyOrderSummary = GetOrderSummaries(mdtFirstDayOfMonth, mdtToday, False, cbytTopSellers)
maryYearlyOrderSummary = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)

pstrProductLink = "<li><prodName>&nbsp;<b>(<salesQty>)</b></li>"
If cblnAddon_ProductMgr Then
	pstrProductLink = "<li><a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodName></a>&nbsp;<b>(<salesQty>)</b></li>"
End If
%>
<table class="tbl" border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" ID="tblOrderSummaries">
    <tr>
    <th rowspan="2">&nbsp;</th>
    <th colspan="4" class="hdrNonSelected">Order Summaries</th>
    </tr>
    <tr class="tdHighlight">
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Day</th>
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Week</th>
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Month</th>
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Year</th>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Orders</th>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryDailyOrderSummary(enOrderSummary_SalesCount) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryWeeklyOrderSummary(enOrderSummary_SalesCount) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryMonthlyOrderSummary(enOrderSummary_SalesCount) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryYearlyOrderSummary(enOrderSummary_SalesCount) %>&nbsp;</td>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Sales</th>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryDailyOrderSummary(enOrderSummary_SalesTotal)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryWeeklyOrderSummary(enOrderSummary_SalesTotal)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryMonthlyOrderSummary(enOrderSummary_SalesTotal)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryYearlyOrderSummary(enOrderSummary_SalesTotal)) %>&nbsp;</td>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Avg. Order</th>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryDailyOrderSummary(enOrderSummary_AverageOrder)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryWeeklyOrderSummary(enOrderSummary_AverageOrder)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryMonthlyOrderSummary(enOrderSummary_AverageOrder)) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= CustomCurrency(maryYearlyOrderSummary(enOrderSummary_AverageOrder)) %>&nbsp;</td>
    </tr>
    <% If cbytTopSellers > -1 Then %>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Best Sellers</th>
    <td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(maryDailyOrderSummary(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
    <td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(maryWeeklyOrderSummary(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
    <td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(maryMonthlyOrderSummary(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
    <td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(maryYearlyOrderSummary(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
    </tr>
    <% End If	'cbytTopSellers > -1 %>
</table>
        <p>
<!--webbot bot="PurpleText" preview="End Order Summaries" --></td>
      </tr>
      <tr>
        <td align="left" valign="top">
<!--webbot bot="PurpleText" preview="Begin Orders Awaiting Payment" -->
<%
If cblnAddon_OrderMgr Then pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=2&optShipment_Filter=0'><orderID></a>"
Call getOrderDetails(maryWorking, "", "", enCriteria_False, enCriteria_DoNotInclude, enCriteria_True)
%>
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblOrdersAwaitingPayment">
      <tr>
        <th colspan="5" class="hdrNonSelected">Orders Awaiting Payment</th>
      </tr>
      <% If isArray(maryWorking) Then %>
      <tr class="tdHighlight">
        <th>Order</th>
        <th>Amount</th>
        <th>Customer</th>
        <th>Tx Type</th>
        <th>Order Date</th>
      </tr>
      <%
      pdblRunningOrderTotal = 0
      pblnEvenRow = False
      For i = 0 To UBound(maryWorking)
		If isNumeric(maryWorking(i)(enOrderDetail_orderGrandTotal)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + maryWorking(i)(enOrderDetail_orderGrandTotal)
      %>
      <tr <%= rowStyle(pblnEvenRow) %>>
        <td align=center><%= Replace(pstrOrderDetailLink, "<orderID>", maryWorking(i)(enOrderDetail_orderID)) %>&nbsp;</td>
        <td align=right><%= CustomCurrency(maryWorking(i)(enOrderDetail_orderGrandTotal)) %>&nbsp;</td>
        <td align=center><%= maryWorking(i)(enOrderDetail_custName) %>&nbsp;</td>
        <td><%= maryWorking(i)(enOrderDetail_orderPaymentMethod) %>&nbsp;</td>
        <td align=center><%= FormatDateTime(maryWorking(i)(enOrderDetail_orderDate), 2) %>&nbsp;</td>
      </tr>
      <% Next 'i %>
      <tr>
        <th><%= UBound(maryWorking)+1 %>&nbsp;</th>
        <th align=right><%= CustomCurrency(pdblRunningOrderTotal) %>&nbsp;</th>
        <th colspan=3>&nbsp;</th>
      </tr>
      <% Else %>
      <tr>
        <th colspan="5">There are no orders with outstanding payments</th>
      </tr>
      <% End If	'isArray(maryWorking) %>
    </table>
        <p>
    <!--webbot bot="PurpleText" preview="End Orders Awaiting Payment" --> </td>
        <td align="center" valign="top">
<!--webbot bot="PurpleText" preview="Begin Orders Awaiting Shipment" -->
<%
If cblnAddon_OrderMgr Then pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=1&optShipment_Filter=1'><orderID></a>"
Call getOrderDetails(maryWorking, "", "", enCriteria_True, enCriteria_False, enCriteria_True)
%>
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblOrdersAwaitingShipment">
      <tr>
        <th colspan="5" class="hdrNonSelected">Orders Awaiting Shipment</th>
      </tr>
      <% If isArray(maryWorking) Then %>
      <tr class="tdHighlight">
        <th>Order</th>
        <th>Amount</th>
        <th>Customer</th>
        <th>Order Date</th>
        <th>Payment Date</th>
      </tr>
      <%
      pdblRunningOrderTotal = 0
      pblnEvenRow = False
      For i = 0 To UBound(maryWorking)
		If isNumeric(maryWorking(i)(enOrderDetail_orderGrandTotal)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + maryWorking(i)(enOrderDetail_orderGrandTotal)
      %>
      <tr <%= rowStyle(pblnEvenRow) %>>
        <td align=center><%= Replace(pstrOrderDetailLink, "<orderID>", maryWorking(i)(enOrderDetail_orderID)) %>&nbsp;</td>
        <td align=right><%= CustomCurrency(maryWorking(i)(enOrderDetail_orderGrandTotal)) %>&nbsp;</td>
        <td align=center><%= maryWorking(i)(enOrderDetail_custName) %>&nbsp;</td>
        <td align=center><%= FormatDateTime(maryWorking(i)(enOrderDetail_orderDate), 2) %>&nbsp;</td>
        <td align=center><%= FormatDateTime(maryWorking(i)(enOrderDetail_orderPaymentDate), 2) %>&nbsp;</td>
      </tr>
      <% Next 'i %>
      <tr>
        <th><%= UBound(maryWorking)+1 %>&nbsp;</th>
        <th align=right><%= CustomCurrency(pdblRunningOrderTotal) %>&nbsp;</th>
        <th colspan=3>&nbsp;</th>
      </tr>
      <% Else %>
      <tr>
        <th colspan="5">There are no orders awaiting shipment</th>
      </tr>
      <% End If	'isArray(maryWorking) %>
    </table>
        <p>
    <!--webbot bot="PurpleText" preview="End Orders Awaiting Shipment" --></td>
      </tr>
      <tr>
        <td align="left" valign="top">
<!--webbot bot="PurpleText" preview="Begin Incomplete Order" -->
<%
If cblnAddon_OrderMgr Then pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=0&optShipment_Filter=0&optIncomplete_Filter=2'><orderID></a>"
Call getOrderDetails(maryWorking, "", "", enCriteria_DoNotInclude, enCriteria_DoNotInclude, enCriteria_False)
%>
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblIncompleteOrders">
      <tr>
        <th colspan="5" class="hdrNonSelected">Incomplete Orders</th>
      </tr>
      <% If isArray(maryWorking) Then %>
      <tr class="tdHighlight">
        <th>Order</th>
        <th>Amount</th>
        <th>Customer</th>
        <th>Tx Type</th>
        <th>Order Date</th>
      </tr>
      <%
      pdblRunningOrderTotal = 0
      pblnEvenRow = False
      For i = 0 To UBound(maryWorking)
		If isNumeric(maryWorking(i)(enOrderDetail_orderGrandTotal)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + maryWorking(i)(enOrderDetail_orderGrandTotal)
      %>
      <tr <%= rowStyle(pblnEvenRow) %>>
        <td align=center><%= Replace(pstrOrderDetailLink, "<orderID>", maryWorking(i)(enOrderDetail_orderID)) %>&nbsp;</td>
        <td align=right><%= CustomCurrency(maryWorking(i)(enOrderDetail_orderGrandTotal)) %>&nbsp;</td>
        <td align=center><%= maryWorking(i)(enOrderDetail_custName) %>&nbsp;</td>
        <td align=center><%= maryWorking(i)(enOrderDetail_orderPaymentMethod) %>&nbsp;</td>
        <td align=center><%= FormatDateTime(maryWorking(i)(enOrderDetail_orderDate), 2) %>&nbsp;</td>
      </tr>
      <% Next 'i %>
      <tr>
        <th><%= UBound(maryWorking)+1 %>&nbsp;</th>
        <th align=right><%= CustomCurrency(pdblRunningOrderTotal) %>&nbsp;</th>
        <th colspan=3>&nbsp;</th>
      </tr>
      <% Else %>
      <tr>
        <th colspan="5">There are no incomplete orders</th>
      </tr>
      <% End If	'isArray(maryWorking) %>
    </table>
        <p>
    <!--webbot bot="PurpleText" preview="End Incomplete Order" --></td>
		<td align="center" valign="top">
		  <h4>Additional Reports</h4>
		  <a href="ssSalesReports.asp?Action=MonthlyReport">Sales Report</a><br>
		  <a href="ssAnnualSalesReports.asp?Action=MonthlyReport">Annual Report</a><br>
		</td>
      </tr>
    </table>
    <p>&nbsp;<!--webbot bot="PurpleText" preview="End Content Section" -->
	<!-- End Content Section -->
    </td>
  </tr>
  <tr>
    <td align="center">
	  <!--webbot bot="PurpleText" preview="Begin Bottom Navigation Section" -->
      <!--#include file="adminFooter.asp"-->
	  <!--webbot bot="PurpleText" preview="End Bottom Navigation Section" -->
     </td>
  </tr>
</table>

<% End Sub %>