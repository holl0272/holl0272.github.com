<% Option Explicit 
'********************************************************************************
'*   Sales Report							                                    *
'*   Release Version:   1.00.003												*
'*   Release Date:		November 15, 2003										*
'*   Revision Date:		October 18, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 1.00.003 (October 18, 2004)										*
'*	   - Restructured code to easily add additional reports						*
'*	   - Added report of detailed product sales									*
'*                                                                              *
'*   Release 1.00.002 (March 7, 2004)											*
'*	   - Added report of sales by category										*
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
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

mstrPageTitle = "Sandshot Software WebStore Manager"

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Const cbytNumProductsToReturn = 0
Const cstrOpenTag = "<ol style='MARGIN-LEFT: 24pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;'>"
Const cstrCloseTag = "</ol>"
	
'/
'/////////////////////////////////////////////////

Dim maryReports(8)
Const enShowReport_Title = 0
Const enShowReport_Display = 1

Const enReportType_BestSellers = 0
Const enReportType_OrderSummary = 1
Const enReportType_OrderSummaryByCategory = 2
Const enReportType_OrderSummaryByCCPayments = 3
Const enReportType_ProductSalesSummary = 4
Const enReportType_ProductSalesSummary_ByCategory = 5
Const enReportType_ProductSalesSummary_ByManufacturer = 6
Const enReportType_CustomerSales = 7
Const enReportType_MostViewedProducts = 8

Call WriteHeader("",True)
Call Main

Response.Flush

'*******************************************************************************************************************************************

Sub setInitialPageDefaults(byRef strStartDate, byRef strStartTime, byRef strEndDate, byRef strEndTime)

	maryReports(enReportType_ProductSalesSummary) = Array("Product Sales Summary", False)
	maryReports(enReportType_ProductSalesSummary_ByCategory) = Array("Product Sales Summary by Category", False)
	maryReports(enReportType_ProductSalesSummary_ByManufacturer) = Array("Product Sales Summary by Manufacturer", False)
	maryReports(enReportType_OrderSummary) = Array("Order Summary", False)
	maryReports(enReportType_OrderSummaryByCategory) = Array("Order Summary By Category", False)
	maryReports(enReportType_OrderSummaryByCCPayments) = Array("Order Summary By CC Card Type", False)
	maryReports(enReportType_BestSellers) = Array("Best Sellers", False)
	maryReports(enReportType_CustomerSales) = Array("Customer Sales", False)
	maryReports(enReportType_MostViewedProducts) = Array("Most Viewed Products", False)
	
	strStartDate = DateAdd("m", -1, Date())
	strStartTime = "12:00:00 AM"
	strEndDate = ""
	strEndTime = ""

End Sub	'setInitialPageDefaults

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

Sub Main

Dim pstrStartDate
Dim pstrStartTime
Dim pstrEndDate
Dim pstrEndTime
Dim i
Dim pblnNeedSpacer
Dim pblnDisplayFilter

	pblnNeedSpacer = False
	pblnDisplayFilter = True

	Call setInitialPageDefaults(pstrStartDate, pstrStartTime, pstrEndDate, pstrEndTime)
	If Len(Request.Form) > 0 Then
		pstrStartDate = LoadRequestValue("StartDate")
		pstrStartTime = LoadRequestValue("StartTime")
		pstrEndDate = LoadRequestValue("EndDate")
		pstrEndTime = LoadRequestValue("EndTime")
		
		For i = 0 To UBound(maryReports)
			maryReports(i)(enShowReport_Display) = CBool(LoadRequestValue("chkShowReport" & i) = "1")
		Next 'i
		pblnDisplayFilter = False
	ElseIf Request.QueryString <> "" Then
		For i = 0 To UBound(maryReports)
			maryReports(i)(enShowReport_Display) = CBool(LoadRequestValue("chkShowReport" & i) = "1")
		Next 'i
		pblnDisplayFilter = False
	End If

%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center" ID="Table2">
  <tr>
    <td width="100%" colspan="2" valign="top">
		<fieldset id="">
			<legend>Report Criteria: <img src="images/filter.bmp" onclick="showHideElement(document.getElementById('tblFilter'));" title="Display Filter"> </legend>
		<form action="ssSalesReports.asp" name="frmData" id="frmData" method="post" style="display:inline">
		<table border="0" cellpadding="2" cellspacing="0" id="tblFilter" style="<%= writeDisplayHide(Not pblnDisplayFilter) %>">
		<tr>
		  <td valign="top">
			<fieldset>
				<legend>Report Type</legend>
				<% For i = 0 To UBound(maryReports) %>
				<input type="checkbox" name="chkShowReport<%= i %>" id="chkShowReport<%= i %>" value="1" <%= isChecked(maryReports(i)(enShowReport_Display)) %>><label for="chkShowReport<%= i %>">&nbsp;<%= maryReports(i)(enShowReport_Title) %></label><br />
				<% Next 'i %>
			</fieldset>
		  </td>
		  <td align="center" valign="top">
			<input class="butn" id="btnFilter" name="btnFilter" type="submit" value="Generate Report"><br />
			<fieldset>
				<legend>Date Range</legend>
				<label for="StartDate">Start Date:&nbsp;</label><input id="StartDate" name="StartDate" size="15" value="<%= pstrStartDate %>">&nbsp;<input id="StartTime" name="StartTime" size="12" value="<%= pstrStartTime %>">
				<a HREF="javascript:doNothing()" title="Select start date"
				onClick="setDateField(document.frmData.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
				<img SRC="images/calendar.gif" BORDER=0></a><br />

				<label for="EndDate">&nbsp;End Date:&nbsp;</label><input id=EndDate name=EndDate size="15" Value="<%= pstrEndDate %>">&nbsp;<input id="EndTime" name="EndTime" size="12" value="<%= pstrEndTime %>">
				<a HREF="javascript:doNothing()" title="Select end date"
				onClick="setDateField(document.frmData.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
				<img SRC="images/calendar.gif" BORDER=0></a>
			</fieldset>
		  </td>
		</tr>
		</table>
		</form>
		</fieldset>
	</td>
  </tr>
<%
'Now for the times
If Len(pstrStartTime) > 0 Then pstrStartDate = Array(pstrStartDate, pstrStartTime)
If Len(pstrEndTime) > 0 Then pstrEndDate = Array(pstrEndDate, pstrEndTime)
		
For i = 0 To UBound(maryReports)
	If maryReports(i)(enShowReport_Display) Then
		If pblnNeedSpacer Then
			Response.Write "<hr />"
		Else
			pblnNeedSpacer = True
		End If
	End If
	
	Select Case i
		Case enReportType_BestSellers: 
			If maryReports(enReportType_BestSellers)(enShowReport_Display) Then
				Call ShowBestSellers(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_OrderSummary: 
			If maryReports(enReportType_OrderSummary)(enShowReport_Display) Then
				Call ShowOrderSummary(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_OrderSummaryByCategory: 
			If maryReports(enReportType_OrderSummaryByCategory)(enShowReport_Display) Then
				Call ShowOrderSummaryByCategory(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_OrderSummaryByCCPayments: 
			If maryReports(enReportType_OrderSummaryByCCPayments)(enShowReport_Display) Then
				Call ShowOrderSummaryByCCPayments(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_ProductSalesSummary: 
			If maryReports(enReportType_ProductSalesSummary)(enShowReport_Display) Then
				Call ShowProductSalesSummary(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_ProductSalesSummary_ByCategory: 
			If maryReports(enReportType_ProductSalesSummary_ByCategory)(enShowReport_Display) Then
				Call ShowProductSalesSummary_ByCategory(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_ProductSalesSummary_ByManufacturer: 
			If maryReports(enReportType_ProductSalesSummary_ByManufacturer)(enShowReport_Display) Then
				Call ShowProductSalesSummary_ByManufacturer(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_CustomerSales: 
			If maryReports(enReportType_CustomerSales)(enShowReport_Display) Then
				Call ShowCustomerSalesSummary(pstrStartDate, pstrEndDate)
			End If
		Case enReportType_MostViewedProducts: 
			If maryReports(enReportType_MostViewedProducts)(enShowReport_Display) Then
				Call ShowMostViewedProductSummary(pstrStartDate, pstrEndDate)
			End If
	End Select
Next 'i
%>
  <tr>
    <td align="center" colspan="2">
	  <!--webbot bot="PurpleText" preview="Begin Bottom Navigation Section" -->
      <!--#include file="adminFooter.asp"-->
	  <!--webbot bot="PurpleText" preview="End Bottom Navigation Section" -->
     </td>
  </tr>
</table>

<% End Sub	'Main %>

<%
Sub ShowBestSellers(byVal strStartDate,byVal strEndDate)

Dim paryWorking
Dim pstrProductLink

	paryWorking = GetOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn)
	pstrProductLink = "<li><prodName>&nbsp;<b>(<salesQty>)</b></li>"
	If cblnAddon_ProductMgr Then pstrProductLink = "<li><a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodName></a>&nbsp;<b>(<salesQty>)</b></li>"
%>
<!--webbot bot="PurpleText" preview="Begin Product Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table1">
    <tr>
    <th class="hdrNonSelected"><%= maryReports(enReportType_BestSellers)(enShowReport_Title) %></th>
    </tr>
	<tr>
	<td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(paryWorking(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
	</tr>
</table>
<!--webbot bot="PurpleText" preview="End Product Summary" -->
<% End Sub	'ShowBestSellers %>

<%
Sub ShowOrderSummary(byVal strStartDate, byVal strEndDate)

Dim paryWorking
Dim pstrOrderDetailLink
Dim pdblRunningOrderTotal
Dim pdblRunningCertificateTotal
Dim pblnEvenRow
Dim i
Dim pblnHasCertificateRedemptions

    pdblRunningOrderTotal = 0
    pdblRunningCertificateTotal = 0
    pblnEvenRow = False
    pblnHasCertificateRedemptions = False
	pstrOrderDetailLink = "<a href='../sfReports1.asp?OrderID=<orderID>'><orderID></a>"
	pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&StartDate=" & DateAdd("d", -0, Date()) & "&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'><orderID></a>"

	Call getOrderDetails(paryWorking, strStartDate, strEndDate, enCriteria_DoNotInclude, enCriteria_DoNotInclude, enCriteria_True)
    If isArray(paryWorking) Then
		For i = 0 To UBound(paryWorking)
			If Not isNumeric(paryWorking(i)(enOrderDetail_orderCertificateRedemptions)) Then paryWorking(i)(enOrderDetail_orderCertificateRedemptions) = 0
			If paryWorking(i)(enOrderDetail_orderCertificateRedemptions) > 0 Then
				pblnHasCertificateRedemptions = True
				Exit For
			End If
		Next 'i
	End If
%>
<!--webbot bot="PurpleText" preview="Begin Order Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblOrderSummary">
    <tr>
    <th colspan="6" class="hdrNonSelected"><%= maryReports(enReportType_OrderSummary)(enShowReport_Title) %></th>
    </tr>
    <% If isArray(paryWorking) Then %>
    <tr class="tdHighlight">
    <th>Order</th>
    <th>Date</th>
    <th>Amount</th>
    <th>Customer</th>
    <th>Tx Type</th>
    <th>Certificates</th>
    </tr>
    <%
    For i = 0 To UBound(paryWorking)
		pdblRunningOrderTotal = pdblRunningOrderTotal + paryWorking(i)(enOrderDetail_orderGrandTotal)
		pdblRunningCertificateTotal = pdblRunningCertificateTotal + paryWorking(i)(enOrderDetail_orderCertificateRedemptions)
    %>
    <tr <%= rowStyle(pblnEvenRow) %>>
    <td align=center><%= Replace(pstrOrderDetailLink, "<orderID>", paryWorking(i)(enOrderDetail_orderID)) %>&nbsp;</td>
    <td align=right><%= FormatDateTime(paryWorking(i)(enOrderDetail_orderDate), 0) %>&nbsp;</td>
    <td align=right><%= WriteCurrency(paryWorking(i)(enOrderDetail_orderGrandTotal)) %>&nbsp;</td>
    <td><%= paryWorking(i)(enOrderDetail_custName) %>&nbsp;</td>
    <td><%= paryWorking(i)(enOrderDetail_orderPaymentMethod) %>&nbsp;</td>
    <td align=right><%= WriteCurrency(paryWorking(i)(enOrderDetail_orderCertificateRedemptions)) %>&nbsp;</td>
    </tr>
    <% Next 'i %>
    <tr>
    <th><%= UBound(paryWorking)+1 %>&nbsp;</th>
    <th>&nbsp;</th>
    <th align=right><%= WriteCurrency(pdblRunningOrderTotal) %>&nbsp;</th>
    <th colspan=2>&nbsp;</th>
    <th align=right><%= WriteCurrency(pdblRunningCertificateTotal) %>&nbsp;</th>
    </tr>
    <% Else %>
    <tr>
    <th colspan="5">No orders meet this criteria</th>
    </tr>
    <% End If	'isArray(paryWorking) %>
</table>
<!--webbot bot="PurpleText" preview="End Order Summary" -->
<% End Sub	'ShowOrderSummary %>

<%
Sub ShowOrderSummaryByCategory(byVal strStartDate, byVal strEndDate)

Dim paryWorking
Dim pstrOrderDetailLink
Dim pdblRunningOrderTotal
Dim pblnEvenRow
Dim i, j
Dim mrsOrders
Dim mlngNumCategories
Dim mstrPrevCategory
Dim mstrCurCategory
Dim maryOutputFields()
Dim maryCategories
Dim mblnHasOrders
Dim mdblGrandTotal
Dim mdblorderHandling
	
Const enSummary_EmptyCategoryName = "-None-"
Const enSummary_Name = 0
Const enSummary_Unit = 1
Const enSummary_Total = 2
Const enSummary_UnitCount = 3

Const enSummary_OrderID = 0
Const enSummary_OrderDate = 1
Const enSummary_OrderTotal = 2
Const enSummary_OrderSubTotal = 3
Const enSummary_Discount = 4
Const enSummary_Shipping = 5
Const enSummary_Tax = 6
Const enSummary_PaymentType = 7
Const enSummary_ShippingMethod = 8
Const enSummary_GC = 9

Const enCategoryStartPos = 10	'highest value above PLUS 1

Dim pblnNewRow
Dim mlngOrderID
Dim mlngOrderID_Prev
Dim mdblSubTotal
Dim mdblTax
Dim mdblShipping
Dim mdblDiscount
Dim mstrPaymentMethod
Dim mdtOrderDate

    pdblRunningOrderTotal = 0
    pblnEvenRow = False
	pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&StartDate=" & DateAdd("d", -0, Date()) & "&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'><orderID></a>"


	Set mrsOrders = GetOrderReport(strStartDate, strEndDate, False)
	mblnHasOrders = CBool(mrsOrders.RecordCount > 0)
	
	'Identify the unique categories
	Dim mdicCategories
	Set mdicCategories = Server.CreateObject("SCRIPTING.DICTIONARY")
	For i = 1 To mrsOrders.RecordCount
		mstrCurCategory = Trim(mrsOrders.Fields("odrdtCategory").Value & "")
		If Len(mstrCurCategory) = 0 Then mstrCurCategory = enSummary_EmptyCategoryName
		If Not mdicCategories.Exists(mstrCurCategory) Then mdicCategories.Add mstrCurCategory, mstrCurCategory
		mrsOrders.MoveNext
	Next 'i

	maryCategories = mdicCategories.Keys
	mlngNumCategories = mdicCategories.Count
	Set mdicCategories = Nothing
	
	'Initialize the output array
	ReDim maryOutputFields(mlngNumCategories+enCategoryStartPos-1)
	maryOutputFields(enSummary_OrderID) = Array("Order #", 0, 0)
	maryOutputFields(enSummary_OrderDate) = Array("Order Date", 0, 0)
	maryOutputFields(enSummary_OrderTotal) = Array("Total Sale", 0, 0)
	maryOutputFields(enSummary_OrderSubTotal) = Array("Subtotal", 0, 0)
	maryOutputFields(enSummary_Shipping) = Array("Shipping", 0, 0)
	maryOutputFields(enSummary_Tax) = Array("Tax", 0, 0)
	maryOutputFields(enSummary_PaymentType) = Array("Payment Type", "", "")
	maryOutputFields(enSummary_ShippingMethod) = Array("Shipping Method", "", "")
	maryOutputFields(enSummary_Discount) = Array("Discount", 0, 0)
	maryOutputFields(enSummary_GC) = Array("Gift Certificates", 0, 0)
	
	'Append the categories to the end
	For i = 0 To UBound(maryCategories)
		maryOutputFields(enCategoryStartPos+i) = Array(maryCategories(i), 0, 0)
	Next 'i
	
	If mblnHasOrders Then mrsOrders.MoveFirst

	%>
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table4">
      <tr>
      <% For i = enSummary_OrderID To enSummary_OrderTotal %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      <% For i = enCategoryStartPos To UBound(maryOutputFields) %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      <% For i = enSummary_OrderSubTotal To enCategoryStartPos - 1 %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      </tr>
      <%
		Do While Not mrsOrders.EOF

			mlngOrderID = mrsOrders.Fields("orderID").Value
			mlngOrderID_Prev = mlngOrderID
			mdtOrderDate = Trim(mrsOrders.Fields("orderDate").Value & "")
			mdblGrandTotal = Trim(mrsOrders.Fields("orderGrandTotal").Value & "")
			If Not isNumeric(mdblGrandTotal) Then mdblGrandTotal = 0
			mdblTax = Trim(mrsOrders.Fields("orderSTax").Value & "")
			If Not isNumeric(mdblTax) Then mdblTax = 0
			mdblShipping = Trim(mrsOrders.Fields("orderShippingAmount").Value & "")
			If Not isNumeric(mdblShipping) Then mdblShipping = 0
			mdblorderHandling = Trim(mrsOrders.Fields("orderHandling").Value & "")
			If Not isNumeric(mdblorderHandling) Then mdblorderHandling = 0
			
			'Now force numbers to be numbers
			mdblGrandTotal = CDbl(mdblGrandTotal)
			mdblTax = CDbl(mdblTax)
			mdblShipping = CDbl(mdblShipping)
			mdblorderHandling = CDbl(mdblorderHandling)
			mdblDiscount = Round(CDbl(mrsOrders.Fields("orderAmount").Value) - mdblGrandTotal + mdblShipping + mdblorderHandling + mdblTax, 2)
			mstrPaymentMethod = Trim(mrsOrders.Fields("orderPaymentMethod").Value & "")
			
			'Set the row totals
			maryOutputFields(enSummary_OrderID)(enSummary_Unit) = mlngOrderID
			maryOutputFields(enSummary_OrderDate)(enSummary_Unit) = mdtOrderDate
			maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit) = mdblGrandTotal - mdblShipping - mdblorderHandling - mdblTax
			maryOutputFields(enSummary_OrderTotal)(enSummary_Unit) = mdblGrandTotal
			maryOutputFields(enSummary_Shipping)(enSummary_Unit) = mdblShipping + mdblorderHandling
			maryOutputFields(enSummary_Tax)(enSummary_Unit) = mdblTax
			maryOutputFields(enSummary_PaymentType)(enSummary_Unit) = mstrPaymentMethod
			maryOutputFields(enSummary_Discount)(enSummary_Unit) = mdblDiscount
			maryOutputFields(enSummary_ShippingMethod)(enSummary_Unit) = Trim(mrsOrders.Fields("orderShipMethod").Value & "")
			maryOutputFields(enSummary_GC)(enSummary_Unit) = Trim(mrsOrders.Fields("SumOfssGCRedemptionAmount").Value & "")
			If Len(maryOutputFields(enSummary_GC)(enSummary_Unit) & "") = 0 Then maryOutputFields(enSummary_GC)(enSummary_Unit) = 0

			'Start a new row
			pblnNewRow = True
			'reset row totals to 0 for categories; other items get set above
			For j = enCategoryStartPos To UBound(maryOutputFields)
				maryOutputFields(j)(enSummary_Unit) = 0
			Next 'j

			Do While pblnNewRow
				mstrCurCategory = Trim(mrsOrders.Fields("odrdtCategory").Value & "")
				If Len(mstrCurCategory) = 0 Then mstrCurCategory = enSummary_EmptyCategoryName
				
				For j = enCategoryStartPos To UBound(maryOutputFields)
					If mstrCurCategory = maryOutputFields(j)(enSummary_Name) Then
						If Len(mrsOrders.Fields("odrdtSubTotal").Value & "")  > 0 And isNumeric(mrsOrders.Fields("odrdtSubTotal").Value) Then maryOutputFields(j)(enSummary_Unit) = maryOutputFields(j)(enSummary_Unit) + CDbl(mrsOrders.Fields("odrdtSubTotal").Value)
						Exit For
					End If
				Next 'j

				If pblnNewRow Then
					mrsOrders.MoveNext
					If mrsOrders.EOF Then
						pblnNewRow = False
					Else
						mlngOrderID = mrsOrders.Fields("orderID").Value
						pblnNewRow = CBool(mlngOrderID = mlngOrderID_Prev)
					End If
				End If
			Loop	'pblnNewRow

			mdblSubTotal = 0
			For j = enCategoryStartPos To UBound(maryOutputFields)
				If maryOutputFields(j)(enSummary_Unit) > 0 Then mdblSubTotal = mdblSubTotal + maryOutputFields(j)(enSummary_Unit)
				'If maryOutputFields(j)(enSummary_Name) = "Bookpacks & Shoulder Bags" Then debugprint " - " & maryOutputFields(j)(enSummary_Name), maryOutputFields(j)(enSummary_Unit)
			Next 'j
			maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit) = mdblSubTotal
			mdblDiscount = Round(maryOutputFields(enSummary_OrderTotal)(enSummary_Unit) - mdblSubTotal - maryOutputFields(enSummary_Shipping)(enSummary_Unit) - maryOutputFields(enSummary_Tax)(enSummary_Unit), 2)
			maryOutputFields(enSummary_Discount)(enSummary_Unit) = mdblDiscount

			'Now output the row			
	  %>
	  <tr>
		<td align="center"><%= Replace(pstrOrderDetailLink, "<orderID>", maryOutputFields(enSummary_OrderID)(enSummary_Unit)) %></td>
		<td align="center"><%= customFormatDateTime(maryOutputFields(enSummary_OrderDate)(enSummary_Unit), 1, "") %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderTotal)(enSummary_Unit)) %></td>
		<%
			For j = enCategoryStartPos To UBound(maryOutputFields)
				If maryOutputFields(j)(enSummary_Unit) > 0 Then
				%>
		<td align="right"><%= WriteCurrency(maryOutputFields(j)(enSummary_Unit)) %></td>
				<%
				Else
					Response.Write "<td>&nbsp;</td>" & vbcrlf
				End If
			Next 'j
		%>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_Discount)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_Shipping)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_Tax)(enSummary_Unit)) %></td>
		<td align="right"><%= maryOutputFields(enSummary_PaymentType)(enSummary_Unit) %></td>
		<td align="right"><%= maryOutputFields(enSummary_ShippingMethod)(enSummary_Unit) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_GC)(enSummary_Unit)) %></td>
	  </tr>
	  <%
	  
			'Add the row totals to the summary totals
			maryOutputFields(enSummary_OrderTotal)(enSummary_Total) = maryOutputFields(enSummary_OrderTotal)(enSummary_Total) + maryOutputFields(enSummary_OrderTotal)(enSummary_Unit)
			maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total) = maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total) + maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)
			maryOutputFields(enSummary_Shipping)(enSummary_Total) = maryOutputFields(enSummary_Shipping)(enSummary_Total) + maryOutputFields(enSummary_Shipping)(enSummary_Unit)
			maryOutputFields(enSummary_Tax)(enSummary_Total) = maryOutputFields(enSummary_Tax)(enSummary_Total) + maryOutputFields(enSummary_Tax)(enSummary_Unit)
			maryOutputFields(enSummary_Discount)(enSummary_Total) = maryOutputFields(enSummary_Discount)(enSummary_Total) + maryOutputFields(enSummary_Discount)(enSummary_Unit)
			maryOutputFields(0)(enSummary_Total) = maryOutputFields(0)(enSummary_Total) + 1
			maryOutputFields(enSummary_GC)(enSummary_Total) = maryOutputFields(enSummary_GC)(enSummary_Total) + maryOutputFields(enSummary_GC)(enSummary_Unit)

			'reset row totals to 0 for categories; other items get set above
			For j = enCategoryStartPos To UBound(maryOutputFields)
				maryOutputFields(j)(enSummary_Total) = maryOutputFields(j)(enSummary_Total) + maryOutputFields(j)(enSummary_Unit)
			Next 'j
	  
		Loop
		Set mrsOrders = Nothing
		
		'Output the footer (summary) row
	  %>
       <tr>
      <% For i = enSummary_OrderID To enSummary_OrderTotal %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      <% For i = enCategoryStartPos To UBound(maryOutputFields) %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      <% For i = enSummary_OrderSubTotal To enCategoryStartPos - 1 %>
        <th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th>
      <% Next 'i %>
      </tr>
      <tr>
		<td class="hdrNonSelected" align="center" colspan="2"><%= maryOutputFields(0)(enSummary_Total) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_OrderTotal)(enSummary_Total)) %></td>
      <% For j = enCategoryStartPos To UBound(maryOutputFields) %>
        <th class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(j)(enSummary_Total)) %></th>
      <% Next 'j %>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)) %></td>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_Discount)(enSummary_Total)) %></td>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_Shipping)(enSummary_Total)) %></td>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_Tax)(enSummary_Total)) %></td>
		<td class="hdrNonSelected">&nbsp;</td>
		<td class="hdrNonSelected">&nbsp;</td>
		<td class="hdrNonSelected"><%= WriteCurrency(maryOutputFields(enSummary_GC)(enSummary_Total)) %></td>
      </tr>
    </table>
     </td>
  </tr>
    <!--webbot bot="PurpleText" preview="End Orders by Category Summary" -->
<% End Sub	'ShowOrderSummaryByCategory %>


<%
Sub ShowOrderSummaryByCCPayments(byVal strStartDate, byVal strEndDate)

Const enSummary_Name = 0
Const enSummary_Unit = 1
Const enSummary_Total = 2
Const enSummary_UnitCount = 3

Const enSummary_OrderID = 0
Const enSummary_OrderDate = 1
Const enSummary_OrderTotal = 2
Const enSummary_OrderSubTotal = 3
Const enSummary_TaxableAmount = 4
Const enSummary_NonTaxableAmount = 5
Const enSummary_Shipping = 6
Const enSummary_Tax = 7
Const enSummary_PaymentType = 8
Const enSummary_GC = 9

Const cblnDisplayCertificateColumn = False
Const cblnSubtractCertifcateFromTotalSale = True

Dim paryWorking
Dim pstrOrderDetailLink
Dim pstrCurTranName
Dim pdblRunningOrderTotal
Dim pblnEvenRow
Dim i, j
Dim mrsOrders
Dim maryOutputFields()
Dim mblnHasOrders
Dim mdblGrandTotal
Dim mdblorderHandling
	
Dim pblnNewRow
Dim pblnNewPaymentMethod
Dim pblnTableOpen
Dim mlngOrderID
Dim mlngOrderID_Prev
Dim pstrPaymentMethod_Prev
Dim mdblTax
Dim mdblShipping
Dim mstrPaymentMethod
Dim mdtOrderDate
Dim pdblCertificate
Dim pdblCheckSales
Dim pdblCheckSubTotal
Dim pdblCheckTotal

    pdblRunningOrderTotal = 0
    pblnEvenRow = False
    pblnTableOpen = False
    
	pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID=<orderID>&Action=ViewOrder&optDisplay=0&StartDate=" & DateAdd("d", -0, Date()) & "&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'><orderID></a>"

	Set mrsOrders = GetOrderReportWithCCDetail(strStartDate, strEndDate, False)
	mblnHasOrders = CBool(mrsOrders.RecordCount > 0)
	
	If mblnHasOrders Then mrsOrders.MoveFirst

	Do While Not mrsOrders.EOF

		mlngOrderID = mrsOrders.Fields("orderID").Value
		mlngOrderID_Prev = mlngOrderID

		mdtOrderDate = Trim(mrsOrders.Fields("orderDate").Value & "")
		mdblGrandTotal = Trim(mrsOrders.Fields("orderGrandTotal").Value & "")
		If Not isNumeric(mdblGrandTotal) Then mdblGrandTotal = 0
		mdblTax = Trim(mrsOrders.Fields("orderSTax").Value & "")
		If Not isNumeric(mdblTax) Then mdblTax = 0
		mdblShipping = Trim(mrsOrders.Fields("orderShippingAmount").Value & "")
		If Not isNumeric(mdblShipping) Then mdblShipping = 0
		mdblorderHandling = Trim(mrsOrders.Fields("orderHandling").Value & "")
		If Not isNumeric(mdblorderHandling) Then mdblorderHandling = 0
		pdblCertificate = Trim(mrsOrders.Fields("SumOfssGCRedemptionAmount").Value & "")
		If Not isNumeric(pdblCertificate) Then pdblCertificate = 0
		
		'Now force numbers to be numbers
		mdblGrandTotal = CDbl(mdblGrandTotal)
		mdblTax = CDbl(mdblTax)
		mdblShipping = CDbl(mdblShipping)
		mdblorderHandling = CDbl(mdblorderHandling)
		
		'mstrPaymentMethod = Trim(mrsOrders.Fields("orderPaymentMethod").Value & "")
		mstrPaymentMethod = Trim(mrsOrders.Fields("transName").Value & "")
		
		pblnNewPaymentMethod = CBool(pstrPaymentMethod_Prev <> mstrPaymentMethod)
		If pblnNewPaymentMethod Then pstrPaymentMethod_Prev = mstrPaymentMethod 
		
		If pblnNewPaymentMethod Then
			If pblnTableOpen Then
	%>
       <tr>
      <% 
      If cblnDisplayCertificateColumn Then
		For i = 0 To UBound(maryOutputFields) 
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      Else
		For i = 0 To UBound(maryOutputFields) - 1
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      End If
      %>
      </tr>
      <tr>
		<td class="hdrNonSelected" align="center" colspan="2"><%= maryOutputFields(0)(enSummary_Total) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderTotal)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_TaxableAmount)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_Shipping)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_Tax)(enSummary_Total)) %></td>
		<td class="hdrNonSelected">
		<%
			'this is to determine if there is a math error
			pdblCheckSubTotal = Round(CDbl(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)), 2)
			pdblCheckSales = Round(CDbl(maryOutputFields(enSummary_TaxableAmount)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total)), 2)

			'taxable sales + non-taxable sales should = subtotal
			If pdblCheckSubTotal = pdblCheckSales Then
				pdblCheckTotal = Round(CDbl(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_Shipping)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_Tax)(enSummary_Total)), 2)
				If cblnSubtractCertifcateFromTotalSale Then pdblCheckTotal = pdblCheckTotal + Round(CDbl(maryOutputFields(enSummary_GC)(enSummary_Total)), 2)
				If Round(CDbl(maryOutputFields(enSummary_OrderTotal)(enSummary_Total)), 2) = pdblCheckTotal Then
					Response.Write "&nbsp;"
				Else
					Response.Write "Error: Total Sale (" & maryOutputFields(enSummary_OrderTotal)(enSummary_Total) & ")does not match sum of subtotal, shipping, and taxes (" & pdblCheckTotal & ")"
				End If
			Else
				Response.Write "Error: Taxable sales + Non-Taxable sales (" & pdblCheckSales & ") do not match order subtotal (" & pdblCheckSubTotal & ")"
			End If
		%>
		</td>
		<% If cblnDisplayCertificateColumn Then %><td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_GC)(enSummary_Total)) %></td><% End If %>
      </tr>
    </table>
	<%
			End If	'pblnTableOpen

		'Initialize the output array
		ReDim maryOutputFields(9)
		maryOutputFields(enSummary_OrderID) = Array("Order #", 0, 0)
		maryOutputFields(enSummary_OrderDate) = Array("Order Date", 0, 0)
		maryOutputFields(enSummary_OrderTotal) = Array("Total Sale", 0, 0)
		maryOutputFields(enSummary_OrderSubTotal) = Array("Subtotal", 0, 0)
		maryOutputFields(enSummary_TaxableAmount) = Array("Taxed Sales", 0, 0)
		maryOutputFields(enSummary_NonTaxableAmount) = Array("Non-Taxed Sales", 0, 0)
		maryOutputFields(enSummary_Shipping) = Array("Shipping", 0, 0)
		maryOutputFields(enSummary_Tax) = Array("IN Tax", 0, 0)
		maryOutputFields(enSummary_PaymentType) = Array("Payment Type", "", "")
		maryOutputFields(enSummary_GC) = Array("Gift Certificates", 0, 0)
	
	%>
	<br />
    <table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
      <tr>
      <% 
      If cblnDisplayCertificateColumn Then
		For i = 0 To UBound(maryOutputFields) 
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      Else
		For i = 0 To UBound(maryOutputFields) - 1
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      End If
      %>
      </tr>
	<%
			pblnTableOpen = True
		End If	'pblnNewPaymentMethod
			
		'Set the row totals
		maryOutputFields(enSummary_OrderID)(enSummary_Unit) = mlngOrderID
		maryOutputFields(enSummary_OrderDate)(enSummary_Unit) = mdtOrderDate
		maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit) = mdblGrandTotal - mdblShipping - mdblorderHandling - mdblTax
		maryOutputFields(enSummary_OrderTotal)(enSummary_Unit) = mdblGrandTotal
		maryOutputFields(enSummary_GC)(enSummary_Unit) = pdblCertificate
		
		'this will subtract out the gc from the total
		If cblnSubtractCertifcateFromTotalSale Then
			maryOutputFields(enSummary_OrderTotal)(enSummary_Unit) = mdblGrandTotal + pdblCertificate
		Else
			maryOutputFields(enSummary_OrderTotal)(enSummary_Unit) = mdblGrandTotal
		End If
		
		If mdblTax > 0 Then
			maryOutputFields(enSummary_TaxableAmount)(enSummary_Unit) = maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)
			maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Unit) = 0
		Else
			maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Unit) = maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)
			maryOutputFields(enSummary_TaxableAmount)(enSummary_Unit) = 0
		End If
		maryOutputFields(enSummary_Shipping)(enSummary_Unit) = mdblShipping + mdblorderHandling
		maryOutputFields(enSummary_Tax)(enSummary_Unit) = mdblTax
		maryOutputFields(enSummary_PaymentType)(enSummary_Unit) = mstrPaymentMethod
		
		'Add the row totals to the summary totals
		maryOutputFields(enSummary_OrderTotal)(enSummary_Total) = maryOutputFields(enSummary_OrderTotal)(enSummary_Total) + maryOutputFields(enSummary_OrderTotal)(enSummary_Unit)
		maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total) = maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total) + maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)
		maryOutputFields(enSummary_TaxableAmount)(enSummary_Total) = maryOutputFields(enSummary_TaxableAmount)(enSummary_Total) + maryOutputFields(enSummary_TaxableAmount)(enSummary_Unit)
		maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total) = maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total) + maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Unit)
		maryOutputFields(enSummary_Shipping)(enSummary_Total) = maryOutputFields(enSummary_Shipping)(enSummary_Total) + maryOutputFields(enSummary_Shipping)(enSummary_Unit)
		maryOutputFields(enSummary_Tax)(enSummary_Total) = maryOutputFields(enSummary_Tax)(enSummary_Total) + maryOutputFields(enSummary_Tax)(enSummary_Unit)
		maryOutputFields(enSummary_GC)(enSummary_Total) = maryOutputFields(enSummary_GC)(enSummary_Total) + maryOutputFields(enSummary_GC)(enSummary_Unit)
		
		maryOutputFields(0)(enSummary_Total) = maryOutputFields(0)(enSummary_Total) + 1

		'Start a new row
		pblnNewRow = True
		'reset row totals to 0 for categories; other items get set above

		Do While pblnNewRow
			pstrCurTranName = Trim(mrsOrders.Fields("transName").Value & "")
			If Len(pstrCurTranName) = 0 Then pstrCurTranName = "-"
			
			For j = 0 To UBound(maryOutputFields)
				If pstrCurTranName = maryOutputFields(j)(enSummary_Name) Then
					If Len(mrsOrders.Fields("odrdtSubTotal").Value & "")  > 0 And isNumeric(mrsOrders.Fields("odrdtSubTotal").Value) Then maryOutputFields(j)(enSummary_Unit) = maryOutputFields(j)(enSummary_Unit) + CDbl(mrsOrders.Fields("odrdtSubTotal").Value)
					Exit For
				End If
			Next 'j

			If pblnNewRow Then
				mrsOrders.MoveNext
				If mrsOrders.EOF Then
					pblnNewRow = False
				Else
					mlngOrderID = mrsOrders.Fields("orderID").Value
					pblnNewRow = CBool(mlngOrderID = mlngOrderID_Prev)
				End If
			End If
		Loop	'pblnNewRow

		'Now output the row			
		'maryOutputFields(j)(enSummary_Total) = maryOutputFields(j)(enSummary_Total) + maryOutputFields(j)(enSummary_Unit)
 			
     %>
		<tr>
		<td align="center"><%= Replace(pstrOrderDetailLink, "<orderID>", maryOutputFields(enSummary_OrderID)(enSummary_Unit)) %></td>
		<td align="center"><%= customFormatDateTime(maryOutputFields(enSummary_OrderDate)(enSummary_Unit), 1, "") %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderTotal)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_TaxableAmount)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_Shipping)(enSummary_Unit)) %></td>
		<td align="right"><%= WriteCurrency(maryOutputFields(enSummary_Tax)(enSummary_Unit)) %></td>
		<td align="right"><%= maryOutputFields(enSummary_PaymentType)(enSummary_Unit) %></td>
		<% If cblnDisplayCertificateColumn Then %><td align="right"><%= WriteCurrency(maryOutputFields(enSummary_GC)(enSummary_Unit)) %></td><% End If %>
	  </tr>
	  <%
		Loop
		Set mrsOrders = Nothing
		
		If pblnTableOpen Then
		'Output the footer (summary) row
	  %>
       <tr>
      <% 
      If cblnDisplayCertificateColumn Then
		For i = 0 To UBound(maryOutputFields) 
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      Else
		For i = 0 To UBound(maryOutputFields) - 1
			%><th class="hdrNonSelected"><%= maryOutputFields(i)(enSummary_Name) %></th><%
		Next 'i
      End If
      %>
      </tr>
      <tr>
		<td class="hdrNonSelected" align="center" colspan="2"><%= maryOutputFields(0)(enSummary_Total) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderTotal)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_TaxableAmount)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_Shipping)(enSummary_Total)) %></td>
		<td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_Tax)(enSummary_Total)) %></td>
		<td class="hdrNonSelected">
		<%
			'this is to determine if there is a math error
			pdblCheckSubTotal = Round(CDbl(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)), 2)
			pdblCheckSales = Round(CDbl(maryOutputFields(enSummary_TaxableAmount)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_NonTaxableAmount)(enSummary_Total)), 2)

			'taxable sales + non-taxable sales should = subtotal
			If pdblCheckSubTotal = pdblCheckSales Then
				pdblCheckTotal = Round(CDbl(maryOutputFields(enSummary_OrderSubTotal)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_Shipping)(enSummary_Total)) + CDbl(maryOutputFields(enSummary_Tax)(enSummary_Total)), 2)
				If cblnSubtractCertifcateFromTotalSale Then pdblCheckTotal = pdblCheckTotal + Round(CDbl(maryOutputFields(enSummary_GC)(enSummary_Total)), 2)
				If Round(CDbl(maryOutputFields(enSummary_OrderTotal)(enSummary_Total)), 2) = pdblCheckTotal Then
					Response.Write "&nbsp;"
				Else
					Response.Write "Error: Total Sale (" & maryOutputFields(enSummary_OrderTotal)(enSummary_Total) & ")does not match sum of subtotal, shipping, and taxes (" & pdblCheckTotal & ")"
				End If
			Else
				Response.Write "Error: Taxable sales + Non-Taxable sales (" & pdblCheckSales & ") do not match order subtotal (" & pdblCheckSubTotal & ")"
			End If
		%>
		</td>
		<% If cblnDisplayCertificateColumn Then %><td class="hdrNonSelected" align="right"><%= WriteCurrency(maryOutputFields(enSummary_GC)(enSummary_Total)) %></td><% End If %>
      </tr>
    </table>
    <%
		End If	'pblnTableOpen
    %>
     </td>
  </tr>
    <!--webbot bot="PurpleText" preview="End Orders by CC Payments Summary" -->
<% End Sub	'ShowOrderSummaryByCCPayments %>

<%
Sub ShowProductSalesSummary(byVal strStartDate, byVal strEndDate)

Dim paryWorking
Dim pstrProductLink
Dim i
Dim pdblRunningOrderTotal
Dim pdblRunningSalesQty
Dim pstrStartDate
Dim pstrEndDate

	pdblRunningOrderTotal = 0
	pdblRunningSalesQty = 0
	'paryWorking = GetOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn)
	paryWorking = GetDetailedOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn, "")
	pstrProductLink = "<prodID>"
	If cblnAddon_ProductMgr Then pstrProductLink = "<a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodID></a>"

	If isArray(strStartDate) Then
		pstrStartDate = strStartDate(0) & " " & strStartDate(1)
	ElseIf Len(strStartDate) = 0 Then
		pstrStartDate = Date()
	Else
		pstrStartDate = strStartDate
	End If
	
	If isArray(strEndDate) Then
		pstrEndDate = strEndDate(0) & " " & strEndDate(1)
	ElseIf Len(strEndDate) = 0 Then
		pstrEndDate = Date()
	Else
		pstrEndDate = strEndDate
	End If
	
%>
<!--webbot bot="PurpleText" preview="Begin Product Sales Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="tblProductSalesSummary">
	<colgroup align="left" />
	<colgroup align="left" />
	<colgroup align="center" />
	<colgroup align="right" />
    <tr>
    <th class="hdrNonSelected" colspan="4" align=center><%= maryReports(enReportType_ProductSalesSummary)(enShowReport_Title) %></th>
    </tr>
    <tr>
    <td colspan="4">Sales Report</td>
    </tr>
    <tr>
    <td colspan="4">Period From:&nbsp;<%= pstrStartDate %>&nbsp;To:&nbsp;<%= pstrEndDate %></td>
    </tr>
    <tr>
		<td class="hdrNonSelected">Product Number</td>
		<td class="hdrNonSelected">Product Name</td>
		<td class="hdrNonSelected">Quantity Sold for the period</td>
		<td class="hdrNonSelected">Revenue</td>
    </tr>
    <% For i = 0 To UBound(paryWorking(enOrderSummary_TopSellers)) %>
    <%
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(0)) Then pdblRunningSalesQty = pdblRunningSalesQty + paryWorking(enOrderSummary_TopSellers)(i)(0)
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(3)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + paryWorking(enOrderSummary_TopSellers)(i)(3)
    %>
    <tr>
		<td><%= Replace(pstrProductLink, "<prodID>", paryWorking(enOrderSummary_TopSellers)(i)(1)) %></td>
		<td><%= writeProductName(paryWorking(enOrderSummary_TopSellers)(i)(2)) %></td>
		<td><%= paryWorking(enOrderSummary_TopSellers)(i)(0) %></td>
		<td><%= WriteCurrency(paryWorking(enOrderSummary_TopSellers)(i)(3)) %></td>
    </tr>
    <% Next 'i %>
    <tr>
		<td class="hdrNonSelected"><%= paryWorking(enOrderSummary_SalesCount) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected"><%= i %> product(s)</td>
		<td class="hdrNonSelected"><%= pdblRunningSalesQty %> item(s)</td>
		<td class="hdrNonSelected"><%= WriteCurrency(pdblRunningOrderTotal) %></td>
    </tr>
</table>
<!--webbot bot="PurpleText" preview="End Product Sales Summary" -->
<% End Sub	'ShowProductSalesSummary %>


<%
Sub ShowProductSalesSummary_ByCategory(byVal strStartDate, byVal strEndDate)

Dim paryWorking
Dim pstrProductLink
Dim i
Dim pdblRunningOrderTotal
Dim pdblRunningSalesQty
Dim pstrStartDate
Dim pstrEndDate

	pdblRunningOrderTotal = 0
	pdblRunningSalesQty = 0
	'paryWorking = GetOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn)
	paryWorking = GetDetailedOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn, "odrdtCategory")

	If isArray(strStartDate) Then
		pstrStartDate = strStartDate(0) & " " & strStartDate(1)
	ElseIf Len(strStartDate) = 0 Then
		pstrStartDate = Date()
	Else
		pstrStartDate = strStartDate
	End If
	
	If isArray(strEndDate) Then
		pstrEndDate = strEndDate(0) & " " & strEndDate(1)
	ElseIf Len(strEndDate) = 0 Then
		pstrEndDate = Date()
	Else
		pstrEndDate = strEndDate
	End If
	
%>
<!--webbot bot="PurpleText" preview="Begin Product Sales Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table6">
	<colgroup align="left" />
	<colgroup align="left" />
	<colgroup align="center" />
	<colgroup align="right" />
    <tr>
    <th class="hdrNonSelected" colspan="4" align=center><%= maryReports(enReportType_ProductSalesSummary_ByCategory)(enShowReport_Title) %></th>
    </tr>
    <tr>
    <td colspan="4">Sales Report</td>
    </tr>
    <tr>
    <td colspan="4">Period From:&nbsp;<%= pstrStartDate %>&nbsp;To:&nbsp;<%= pstrEndDate %></td>
    </tr>
    <tr>
		<td class="hdrNonSelected">Product Number</td>
		<td class="hdrNonSelected">Product Name</td>
		<td class="hdrNonSelected">Quantity Sold for the period</td>
		<td class="hdrNonSelected">Revenue</td>
    </tr>
    <% For i = 0 To UBound(paryWorking(enOrderSummary_TopSellers)) %>
    <%
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(0)) Then pdblRunningSalesQty = pdblRunningSalesQty + paryWorking(enOrderSummary_TopSellers)(i)(0)
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(3)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + paryWorking(enOrderSummary_TopSellers)(i)(3)
    %>
    <tr>
		<td><%= Replace(pstrProductLink, "<prodID>", paryWorking(enOrderSummary_TopSellers)(i)(1)) %></td>
		<td><%= writeProductName(paryWorking(enOrderSummary_TopSellers)(i)(2)) %></td>
		<td><%= paryWorking(enOrderSummary_TopSellers)(i)(0) %></td>
		<td><%= WriteCurrency(paryWorking(enOrderSummary_TopSellers)(i)(3)) %></td>
    </tr>
    <% Next 'i %>
    <tr>
		<td class="hdrNonSelected"><%= paryWorking(enOrderSummary_SalesCount) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected"><%= i %> product(s)</td>
		<td class="hdrNonSelected"><%= pdblRunningSalesQty %> item(s)</td>
		<td class="hdrNonSelected"><%= WriteCurrency(pdblRunningOrderTotal) %></td>
    </tr>
</table>
<!--webbot bot="PurpleText" preview="End Product Sales Summary" -->
<% End Sub	'ShowProductSalesSummary_ByCategory %>


<%
Sub ShowProductSalesSummary_ByManufacturer(byVal strStartDate, byVal strEndDate)

Dim paryWorking
Dim pstrProductLink
Dim i
Dim pdblRunningOrderTotal
Dim pdblRunningSalesQty
Dim pstrStartDate
Dim pstrEndDate

	pdblRunningOrderTotal = 0
	pdblRunningSalesQty = 0
	'paryWorking = GetOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn)
	paryWorking = GetDetailedOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn, "odrdtManufacturer")
	pstrProductLink = "<prodID>"
	If cblnAddon_ProductMgr Then pstrProductLink = "<a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodID></a>"

	If isArray(strStartDate) Then
		pstrStartDate = strStartDate(0) & " " & strStartDate(1)
	ElseIf Len(strStartDate) = 0 Then
		pstrStartDate = Date()
	Else
		pstrStartDate = strStartDate
	End If
	
	If isArray(strEndDate) Then
		pstrEndDate = strEndDate(0) & " " & strEndDate(1)
	ElseIf Len(strEndDate) = 0 Then
		pstrEndDate = Date()
	Else
		pstrEndDate = strEndDate
	End If
	
%>
<!--webbot bot="PurpleText" preview="Begin Product Sales Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table7">
	<colgroup align="left" />
	<colgroup align="left" />
	<colgroup align="center" />
	<colgroup align="right" />
    <tr>
    <th class="hdrNonSelected" colspan="4" align=center><%= maryReports(enReportType_ProductSalesSummary_ByManufacturer)(enShowReport_Title) %></th>
    </tr>
    <tr>
    <td colspan="4">Sales Report</td>
    </tr>
    <tr>
    <td colspan="4">Period From:&nbsp;<%= pstrStartDate %>&nbsp;To:&nbsp;<%= pstrEndDate %></td>
    </tr>
    <tr>
		<td class="hdrNonSelected">Product Number</td>
		<td class="hdrNonSelected">Product Name</td>
		<td class="hdrNonSelected">Quantity Sold for the period</td>
		<td class="hdrNonSelected">Revenue</td>
    </tr>
    <% For i = 0 To UBound(paryWorking(enOrderSummary_TopSellers)) %>
    <%
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(0)) Then pdblRunningSalesQty = pdblRunningSalesQty + paryWorking(enOrderSummary_TopSellers)(i)(0)
		If isNumeric(paryWorking(enOrderSummary_TopSellers)(i)(3)) Then pdblRunningOrderTotal = pdblRunningOrderTotal + paryWorking(enOrderSummary_TopSellers)(i)(3)
    %>
    <tr>
		<td><%= Replace(pstrProductLink, "<prodID>", paryWorking(enOrderSummary_TopSellers)(i)(1)) %></td>
		<td><%= writeProductName(paryWorking(enOrderSummary_TopSellers)(i)(2)) %></td>
		<td><%= paryWorking(enOrderSummary_TopSellers)(i)(0) %></td>
		<td><%= WriteCurrency(paryWorking(enOrderSummary_TopSellers)(i)(3)) %></td>
    </tr>
    <% Next 'i %>
    <tr>
		<td class="hdrNonSelected"><%= paryWorking(enOrderSummary_SalesCount) %> order(s)&nbsp;</td>
		<td class="hdrNonSelected"><%= i %> product(s)</td>
		<td class="hdrNonSelected"><%= pdblRunningSalesQty %> item(s)</td>
		<td class="hdrNonSelected"><%= WriteCurrency(pdblRunningOrderTotal) %></td>
    </tr>
</table>
<!--webbot bot="PurpleText" preview="End Product Sales Summary" -->
<% End Sub	'ShowProductSalesSummary_ByManufacturer %>

<%
Sub ShowCustomerSalesSummary(byVal strStartDate, byVal strEndDate)

Dim pstrCustomerLink
Dim pstrCustomerName
Dim i
Dim pstrStartDate
Dim pstrEndDate
Dim pstrSQL
Dim pstrsqlWhere
Dim pobjRS

	pstrCustomerLink = "<a href='sfCustomerAdmin.asp?Action=viewItem&ViewID=<custID>'><custName></a>"

	If isArray(strStartDate) Then
		pstrStartDate = strStartDate(0) & " " & strStartDate(1)
	ElseIf Len(strStartDate) = 0 Then
		pstrStartDate = Date()
	Else
		pstrStartDate = strStartDate
	End If
	
	If isArray(strEndDate) Then
		pstrEndDate = strEndDate(0) & " " & strEndDate(1)
	ElseIf Len(strEndDate) = 0 Then
		pstrEndDate = Date()
	Else
		pstrEndDate = strEndDate
	End If
	
	pstrsqlWhere = " Where (sfOrders.orderIsComplete=1)"
	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " And (sfOrders.orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " And (sfOrders.orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"

	If cblnSQLDatabase Then
		pstrSQL = "SELECT sfCustomers.custID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custIsSubscribed, sfCustomers.custTimesAccessed, sfCustomers.custLastAccess, PricingLevels.PricingLevelName, Sum(convert(money,sfOrders.orderAmount)) AS SumOforderAmount, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity,  Max(sfOrders.orderDate) AS MaxOforderDate" _
				& " FROM ((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN PricingLevels ON sfCustomers.PricingLevelID = PricingLevels.PricingLevelID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
				& pstrsqlWhere _
				& " GROUP BY sfCustomers.custID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custIsSubscribed, sfCustomers.custTimesAccessed, sfCustomers.custLastAccess, PricingLevels.PricingLevelName" _
				& " HAVING (sfCustomers.custID Is Not Null)" _
				& " ORDER BY Sum(convert(money,sfOrders.orderAmount)) DESC , Sum(sfOrderDetails.odrdtQuantity) DESC"
	Else
		pstrSQL = "SELECT sfCustomers.custID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custIsSubscribed, sfCustomers.custTimesAccessed, sfCustomers.custLastAccess, PricingLevels.PricingLevelName, Sum(sfOrders.orderAmount) AS SumOforderAmount, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity,  Max(sfOrders.orderDate) AS MaxOforderDate" _
				& " FROM ((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN PricingLevels ON sfCustomers.PricingLevelID = PricingLevels.PricingLevelID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
				& pstrsqlWhere _
				& " GROUP BY sfCustomers.custID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custIsSubscribed, sfCustomers.custTimesAccessed, sfCustomers.custLastAccess, PricingLevels.PricingLevelName" _
				& " HAVING (sfCustomers.custID Is Not Null)" _
				& " ORDER BY Sum(sfOrders.orderAmount) DESC , Sum(sfOrderDetails.odrdtQuantity) DESC"
	End If
	
	Set pobjRS = GetRS(pstrSQL)
	With pobjRS
%>
<!--webbot bot="PurpleText" preview="Begin Customer Sales Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table3">
	<colgroup align="left" />
	<colgroup align="center" />
	<colgroup align="center" />
	<colgroup align="right" />
	<colgroup align="center" />
	<colgroup align="center" />
	<colgroup align="right" />
	<colgroup align="right" />
    <tr>
    <th class="hdrNonSelected" colspan="8" align=center><%= maryReports(enReportType_CustomerSales)(enShowReport_Title) %></th>
    </tr>
    <tr>
    <td colspan="8">Sales Report</td>
    </tr>
    <tr>
    <td colspan="8">Period From:&nbsp;<%= pstrStartDate %>&nbsp;To:&nbsp;<%= pstrEndDate %></td>
    </tr>
    <tr>
		<td class="hdrNonSelected" align="left">Name</td>
		<td class="hdrNonSelected" align="center">Subscribed</td>
		<td class="hdrNonSelected" align="center">Pricing Level</td>
		<td class="hdrNonSelected" align="center">Sales</td>
		<td class="hdrNonSelected" align="center">Item Count</td>
		<td class="hdrNonSelected" align="center">Visits</td>
		<td class="hdrNonSelected" align="center">Last Visited</td>
		<td class="hdrNonSelected" align="center">Last Ordered</td>
    </tr>
    <% Do While Not .EOF %>
    <%
		pstrCustomerName = .Fields("custLastName").Value & ", " & .Fields("custFirstName").Value
		If Len(.Fields("custMiddleInitial").Value) > 0 Then pstrCustomerName = pstrCustomerName & " " & .Fields("custMiddleInitial").Value
    %>
    <tr>
		<td><%= Replace(Replace(pstrCustomerLink, "<custName>", pstrCustomerName), "<custID>", .Fields("custID").Value) %></td>
		<td><%= ConvertToBoolean(.Fields("custIsSubscribed").Value, False) %></td>
		<td><%= .Fields("PricingLevelName").Value %></td>
		<td>&nbsp;&nbsp;<%= WriteCurrency(.Fields("SumOforderAmount").Value) %>&nbsp;&nbsp;</td>
		<td><%= .Fields("SumOfodrdtQuantity").Value %></td>
		<td><%= .Fields("custTimesAccessed").Value %></td>
		<td><%= customFormatDateTime(.Fields("custLastAccess").Value, 1, "") %></td>
		<td><%= customFormatDateTime(.Fields("MaxOforderDate").Value, 1, "") %></td>
    </tr>
    <%	.MoveNext %>
    <% Loop %>
</table>
<!--webbot bot="PurpleText" preview="End Customer Sales Summary" -->
<%
	End With
	Call ReleaseObject(pobjRS)
%>
<% End Sub	'ShowCustomerSalesSummary %>

<%
Sub ShowMostViewedProductSummary(byVal strStartDate, byVal strEndDate)

Dim pstrProductLink
Dim pstrProductName
Dim i
Dim pstrStartDate
Dim pstrEndDate
Dim pstrSQL
Dim pstrsqlWhere
Dim pobjRS
Dim pdicProducts
Dim pstrProdID
Dim paryTemp
Dim paryItem
Dim vItem
Dim paryWorking

	pstrProductLink = "<a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodName></a>"

	If isArray(strStartDate) Then
		pstrStartDate = strStartDate(0) & " " & strStartDate(1)
	ElseIf Len(strStartDate) = 0 Then
		pstrStartDate = Date()
	Else
		pstrStartDate = strStartDate
	End If
	
	If isArray(strEndDate) Then
		pstrEndDate = strEndDate(0) & " " & strEndDate(1)
	ElseIf Len(strEndDate) = 0 Then
		pstrEndDate = Date()
	Else
		pstrEndDate = strEndDate
	End If
	
	'Load the active products
	Set pdicProducts = Server.CreateObject("Scripting.Dictionary")
	pstrSQL = "SELECT prodID, prodName From sfProducts Where prodEnabledIsActive=1 Order By prodName"
	Set pobjRS = GetRS(pstrSQL)
	With pobjRS
		Do While Not .EOF
			pstrProdID = Trim(.Fields("prodID").Value & "")
			If Not pdicProducts.Exists(pstrProdID) Then pdicProducts.Add pstrProdID, Array(Trim(.Fields("prodName").Value & ""), 0, "", 0, "")
			.MoveNext
		Loop
	End With
	Call ReleaseObject(pobjRS)

	'Load the sales
	paryWorking = GetOrderSummaries(strStartDate, strEndDate, False, cbytNumProductsToReturn)
	If isArray(paryWorking) Then
		If isArray(paryWorking(enOrderSummary_TopSellers)) Then
			For i = 0 To UBound(paryWorking(enOrderSummary_TopSellers))
				pstrProdID = paryWorking(enOrderSummary_TopSellers)(i)(1)
				If pdicProducts.Exists(pstrProdID) Then
					paryItem = pdicProducts(pstrProdID)
					paryItem(3) = paryItem(3) + paryWorking(enOrderSummary_TopSellers)(i)(0)
					pdicProducts(pstrProdID) = paryItem
				End If
			Next 'i
		End If	'isArray(paryWorking(enOrderSummary_TopSellers))
	End If	'isArray(paryWorking)

	'Load the product views
	pstrSQL = "SELECT visitorRecentlyViewedProducts, visitorLastVisited FROM visitors" & pstrsqlWhere
	pstrsqlWhere = " Where (visitorRecentlyViewedProducts Is Not Null)"
	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " And (visitorLastVisited >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " And (visitorLastVisited <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"

	Set pobjRS = GetRS(pstrSQL & pstrsqlWhere)
	With pobjRS
		Do While Not .EOF
			paryTemp = Split(.Fields("visitorRecentlyViewedProducts").Value, "|")
			For i = 0 To UBound(paryTemp)
				pstrProdID = paryTemp(i)
				If pdicProducts.Exists(pstrProdID) Then
					paryItem = pdicProducts(pstrProdID)
					paryItem(1) = paryItem(1) + 1
					If .Fields("visitorLastVisited").Value < paryItem(2) Then paryItem(2) = .Fields("visitorLastVisited").Value
					pdicProducts(pstrProdID) = paryItem
				End If
			Next
			.MoveNext
		Loop
	End With
	Call ReleaseObject(pobjRS)
	
	Dim pobjxmlDoc
	Dim pobjxmlRoot
	Dim pobjxmlNode
	Dim pobjxmlElement
	
	set pobjxmlDoc = server.CreateObject("MSXML2.DOMDocument")
	
	' Create processing instruction and document root
    Set pobjxmlNode = pobjxmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set pobjxmlNode = pobjxmlDoc.insertBefore(pobjxmlNode, pobjxmlDoc.childNodes.Item(0))
    
	' Create document root
    Set pobjxmlRoot = pobjxmlDoc.createElement("Products")
    Set pobjxmlDoc.documentElement = pobjxmlRoot
    pobjxmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"

    For Each vItem In pdicProducts
		paryTemp = pdicProducts(vItem)
		pstrProductName = paryTemp(0)
		
		'Create the product node
		Set pobjxmlNode = pobjxmlDoc.createElement("Product")
		pobjxmlRoot.appendChild pobjxmlNode

		Set pobjxmlElement = pobjxmlDoc.createElement("ProductCode")
		pobjxmlElement.Text = vItem
		pobjxmlNode.appendChild pobjxmlElement

		Set pobjxmlElement = pobjxmlDoc.createElement("ProductName")
		pobjxmlElement.Text = pstrProductName
		pobjxmlNode.appendChild pobjxmlElement

		Set pobjxmlElement = pobjxmlDoc.createElement("ProductViews")
		pobjxmlElement.Text = paryTemp(1)
		pobjxmlNode.appendChild pobjxmlElement

		Set pobjxmlElement = pobjxmlDoc.createElement("ProductSales")
		pobjxmlElement.Text = paryTemp(3)
		pobjxmlNode.appendChild pobjxmlElement

		Set pobjxmlElement = pobjxmlDoc.createElement("ProductLastViewed")
		pobjxmlElement.Text = customFormatDateTime(paryTemp(2), 1, "")
		pobjxmlNode.appendChild pobjxmlElement

	Next
	
	'Response.Write "<fieldset><legend>XML</legend>" & Server.HTMLEncode(pobjXMLDoc.xml) & "</fieldset>"
    Set pobjxmlElement = Nothing
    Set pobjxmlNode = Nothing
    Set pobjxmlRoot = Nothing
    Set pobjxmlDoc = Nothing
	
	
%>
<!--webbot bot="PurpleText" preview="Begin Most Viewed Product Summary" -->
<table class="tbl" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" id="Table5">
	<colgroup align="left" />
	<colgroup align="center" />
	<colgroup align="center" />
	<colgroup align="center" />
    <tr>
    <th class="hdrNonSelected" colspan="4" align=center><%= maryReports(enReportType_MostViewedProducts)(enShowReport_Title) %></th>
    </tr>
    <tr>
    <td colspan="4">Sales Report</td>
    </tr>
    <tr>
    <td colspan="4">Period From:&nbsp;<%= pstrStartDate %>&nbsp;To:&nbsp;<%= pstrEndDate %></td>
    </tr>
    <tr>
		<td class="hdrNonSelected" align="left">Product</td>
		<td class="hdrNonSelected" align="center">Views</td>
		<td class="hdrNonSelected" align="center">Sales</td>
		<td class="hdrNonSelected" align="center">Last View</td>
    </tr>
    <% For Each vItem In pdicProducts %>
    <%
		paryTemp = pdicProducts(vItem)
		pstrProductName = paryTemp(0)
    %>
    <tr>
		<td><%= pstrProductName %></td>
		<td><%= paryTemp(1) %></td>
		<td><%= paryTemp(3) %></td>
		<td><%= customFormatDateTime(paryTemp(2), 1, "") %></td>
    </tr>
    <% Next %>
</table>
<!--webbot bot="PurpleText" preview="End Most Viewed Product Summary" -->
<%
	Call ReleaseObject(pdicProducts)
%>
<% End Sub	'ShowMostViewedProductSummary %>
