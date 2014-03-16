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
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<%

	mstrPageTitle = "Sandshot Software WebStore Manager"
	Call CheckLoginStatus_AdminPage(False)
%><!--#include file="adminFooter.asp"--><%
	If Response.Buffer Then Response.Flush
    Call ReleaseObject(cnn)

'*******************************************************************************************************************************************

Function WriteDifference(byVal dblNew, byVal dblOld, byVal blnFormatCurrency, byVal blnShowDifference) 

Dim pdblDifference
Dim pdblRatio
Dim pstrOut_New
Dim pstrOut_Old
Dim pstrOut_Ratio
Dim pstrOut

	pdblDifference = CDbl(dblNew) - CDbl(dblOld)
	If pdblDifference < 0.0001 Then pdblDifference = 0
	If dblOld = 0 Then
		pdblRatio = 0
		pstrOut_Ratio = ""
	Else
		pdblRatio = (dblNew - dblOld)/dblOld
		pstrOut_Ratio = FormatPercent(pdblRatio, 2)
	End If

	If blnFormatCurrency Then
		pstrOut_New = FormatCurrency(dblNew)
		pstrOut_Old = FormatCurrency(dblOld)
	Else
		pstrOut_New = dblNew
		pstrOut_Old = dblOld
	End If
	
	pstrOut = pstrOut_New
	If blnShowDifference Then
		If pdblDifference > 0 Then
			pstrOut = pstrOut & "<span style=""font-family: Wingdings;color:green"" title=""" & pstrOut_Old & " (+" & pstrOut_Ratio & ")"">ñ</span>"
		ElseIf pdblDifference < 0 Then
			pstrOut = pstrOut & "<span style=""font-family: Wingdings;color:red"" title=""" & pstrOut_Old & " (-" & pstrOut_Ratio & ")"">ò</span>"
		Else
			pstrOut = pstrOut & "<span style=""font-family: Wingdings;color:black"" title=""No Change"">ó</span>"
		End If
	Else
		If pdblDifference > 0 Then
			pstrOut = "<span style=""color:green"" title=""" & pstrOut_Old & " (+" & pstrOut_Ratio & ")"">" & pstrOut & "</span>"
		ElseIf pdblDifference < 0 Then
			pstrOut = "<span style=""color:red"" title=""" & pstrOut_Old & " (-" & pstrOut_Ratio & ")"">" & pstrOut & "</span>"
		Else
			pstrOut = "<span style=""color:black"" title=""No Change"">" & pstrOut & "</span>"
		End If
	End If

	WriteDifference = pstrOut
	
End Function

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
pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&StartDate=" & DateAdd("d", -0, Date()) & "&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'><orderID></a>"
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
Dim maryDailyOrderSummary_PriorPeriod
Dim maryWeeklyOrderSummary_PriorPeriod
Dim maryMonthlyOrderSummary_PriorPeriod
Dim maryYearlyOrderSummary_PriorPeriod

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
Const cblnShowArrows = True
	
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

maryDailyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("d", -1, mdtToday), DateAdd("d", -1, mdtToday), False, -1)
maryWeeklyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("d", -7, mdtFirstDayOfWeek), DateAdd("d", -7, mdtToday), False, -1)
maryMonthlyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("m", -1, mdtFirstDayOfMonth), DateAdd("m", -1, mdtToday), False, -1)
maryYearlyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("y", -1, mdtFirstDayOfYear), DateAdd("y", -1, mdtToday), False, -1)

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
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_SalesCount), maryDailyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_SalesCount), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_SalesCount), maryMonthlyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_SalesCount), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %>&nbsp;</td>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Sales</th>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_SalesTotal), maryDailyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_SalesTotal), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_SalesTotal), maryMonthlyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_SalesTotal), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %>&nbsp;</td>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Avg. Order</th>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_AverageOrder), maryDailyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_AverageOrder), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_AverageOrder), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), False, cblnShowArrows) %>&nbsp;</td>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_AverageOrder), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), False, cblnShowArrows) %>&nbsp;</td>
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
pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=2&optShipment_Filter=0'><orderID></a>"
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
pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=1&optShipment_Filter=1'><orderID></a>"
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
pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&optPayment_Filter=0&optShipment_Filter=0&optIncomplete_Filter=2'><orderID></a>"
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
		  <a href="ssSalesReports.asp?Action=MonthlyReport">Sales Report</a><br />
		  <a href="ssAnnualSalesReports.asp?Action=MonthlyReport">Annual Report</a><br />
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