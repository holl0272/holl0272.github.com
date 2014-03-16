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
Server.ScriptTimeout = 600			'in seconds. Adjust for large databases or if some products have a lot of attributes. Server Default is usually 90 seconds
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

mstrPageTitle = "Sandshot Software WebStore Manager"

	Call WriteHeader("",True)
	Call Main

	If Response.Buffer Then Response.Flush

'******************************************************************************************************************************************************************

Function rowStyle(byRef blnEvenRow)
	If blnEvenRow Then
		rowStyle = "class='Inactive'"
	Else
		rowStyle = ""
	End If
	blnEvenRow = Not blnEvenRow
End Function	'rowStyle

'******************************************************************************************************************************************************************

Function firstSaleYear()

Dim pstrResult
Dim pstrSQL

	If cblnSQLDatabase Then
		pstrSQL = "Select Top 1 orderDate From sfOrders Where orderDate is not Null Order By convert(datetime,orderDate)"
	Else
		pstrSQL = "Select Top 1 orderDate From sfOrders Where orderDate is not Null Order By CDate(orderDate)"
	End If

	pstrResult = getReturnValue(pstrSQL)
	
	If isDate(pstrResult) Then
		pstrResult = Year(pstrResult)
	Else
		pstrResult = "2004"
	End If
	
	firstSaleYear = pstrResult
	
End Function	'firstSaleYear

'*******************************************************************************************************************************************

Sub Main

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Const cbytTopSellers = 0
Const cstrOpenTag = "<ol style='MARGIN-LEFT: 18pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;'>"
Const cstrCloseTag = "</ol>"
	
'/
'/////////////////////////////////////////////////

Dim pstrOrderDetailLink
Dim pdblRunningOrderTotal
Dim i, j
Dim maryWorking
Dim pblnEvenRow
Dim pstrAction
Dim maryOrderSummary()
Dim maryOrderSummary_PriorPeriod()
Dim maryOrderSummaryHeaders
Dim mstrYear
Dim mdtToday
Dim mdtFirstDay
Dim mdtFirstDayOfWeek
Dim mdtFirstDayOfMonth
Dim mdtFirstDayOfYear
Dim pstrProductLink
Dim pstrStartYear
Dim pstrTableCaption

Dim clngStartYear
Dim pblnShowByMonth
Dim plngCurrentYear
Dim paryYears
Dim paryYearsSelected
Dim mlngNumYearsSelected

	clngStartYear = firstSaleYear
	plngCurrentYear = Year(Date())

	ReDim paryYears(plngCurrentYear - clngStartYear)
	For i = 0 To UBound(paryYears)
		paryYears(i) = clngStartYear + UBound(paryYears) - i
	Next 'i

	ReDim paryYearsSelected(UBound(paryYears))
	mlngNumYearsSelected = -1
	For i = 0 To UBound(paryYearsSelected)
		paryYearsSelected(i) = CBool(Request.QueryString("Year" & paryYears(i)) = "1")
		If paryYearsSelected(i) Then mlngNumYearsSelected = mlngNumYearsSelected + 1
	Next 'i
	
	pstrAction = Request.QueryString("Action")
	pstrStartYear = Request.QueryString("StartYear")
	pblnShowByMonth = CBool(Request.QueryString("ShowByMonth") = "1")
	
	'Link to sfReports
	pstrOrderDetailLink = "<a href='../sfReports1.asp?OrderID=<orderID>'><orderID></a>"
	pblnEvenRow = False
%>
<script language="javascript">
	function viewAnnualReport(theSelect)
	{
		var queryString = theSelect.options[theSelect.selectedIndex].value;
		
		//alert(queryString);
		//alert(window.location.href);
		window.location.href = "ssAnnualSalesReports.asp?" + queryString;
	}
</script>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" valign="top" align="left">
	<!-- Begin Content Section -->
	<form name="frmData" id="frmData" action="ssAnnualSalesReports.asp" method="get">
	  <input type="hidden" name="Action" id="Action" value="<%= pstrAction %>">
    <p><label for="salesReports" title="">Select a report to view</label>:&nbsp;
    <select onchange="viewAnnualReport(this);" id="salesReports">
    <option value="" <%= isSelected(ConvertToBoolean(pstrStartYear = "", False)) %>>Current Summary</option>
    <% For i = 0 To UBound(paryYears) %>
		<option value="Action=MonthlyReport&StartYear=<%= paryYears(i) %>" <%= isSelected(ConvertToBoolean(CStr(pstrStartYear) = CStr(paryYears(i)), False)) %>>Sales By Month for <%= paryYears(i) %></option>
    <% Next 'i %>
    <option value="Action=YearlyReport" <%= isSelected(ConvertToBoolean(pstrStartYear = "" And pstrAction = "YearlyReport", False)) %>>Yearly Reports</option>
    </select>
    <% If pstrAction = "YearlyReport" Then %>
    <hr />
    <% For i = 0 To UBound(paryYears) %>
		<input type="checkbox" name="Year<%= paryYears(i) %>" id="Year<%= paryYears(i) %>" value="1" <%= isChecked(paryYearsSelected(i)) %>>&nbsp;<label for="Year<%= paryYears(i) %>"><%= paryYears(i) %></label> 
    <% Next 'i %>
    <br />
    <input type="checkbox" name="ShowByMonth" id="ShowByMonth" value="1" <%= isChecked(pblnShowByMonth) %>>&nbsp;<label for="ShowByMonth">Show by Month</label>
    <input class="butn" type="submit" name="btnAction" id="btnAction" value="Submit">
    <% End If %>
     </p>
	</form>

<!--webbot bot="PurpleText" preview="Begin Order Reports" -->
<%

mdtToday = Date()
If Len(pstrStartYear) = 0 Then
	mstrYear = CStr(Year(mdtToday))
Else
	mstrYear = pstrStartYear
End If

If True Then
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

Select Case pstrAction
	Case "YearlyReport":
		ReDim maryOrderSummaryHeaders(mlngNumYearsSelected)

		mlngNumYearsSelected = -1
		For i = 0 To UBound(paryYearsSelected)
			If paryYearsSelected(i) Then
				mlngNumYearsSelected = mlngNumYearsSelected + 1
				maryOrderSummaryHeaders(mlngNumYearsSelected) = paryYears(i)
				If Len(pstrTableCaption) = 0 Then
					pstrTableCaption = " for " & paryYears(i)
				Else
					pstrTableCaption = pstrTableCaption & ", " & paryYears(i)
				End If
			End If
		Next 'i

		If pblnShowByMonth Then
			ReDim maryOrderSummary((mlngNumYearsSelected + 1) * 12 - 1)
			mlngNumYearsSelected = 0
			For i = 0 To UBound(paryYearsSelected)
				If paryYearsSelected(i) Then
					mdtFirstDayOfYear = CDate(CStr(paryYears(i)) + "/1/1")

					For j = 0 To 11
						mdtFirstDayOfMonth = DateAdd("m", j, mdtFirstDayOfYear)
						mdtToday = DateAdd("d", -1, DateAdd("m", 1, mdtFirstDayOfMonth))

						maryOrderSummary(mlngNumYearsSelected) = GetOrderSummaries(mdtFirstDayOfMonth, mdtToday, False, cbytTopSellers)
						mlngNumYearsSelected = mlngNumYearsSelected + 1
					Next 'i
					mdtToday = CDate(CStr(paryYears(i)) + "/12/31")
					'maryOrderSummary(mlngNumYearsSelected) = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)
				End If
			Next 'i
		Else
			ReDim maryOrderSummary(mlngNumYearsSelected)
			mlngNumYearsSelected = 0
			For i = 0 To UBound(paryYearsSelected)
				If paryYearsSelected(i) Then
					mdtFirstDayOfYear = CDate(CStr(paryYears(i)) + "/1/1")
					mdtToday = CDate(CStr(paryYears(i)) + "/12/31")

					maryOrderSummary(mlngNumYearsSelected) = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)
					mlngNumYearsSelected = mlngNumYearsSelected + 1
				End If
			Next 'i
		End If

		'Now build the summary
		Dim maryProducts
		maryProducts = createProductSalesReport(maryOrderSummary)
		
		'maryOrderSummary(12) = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)
	Case "MonthlyReport":
		pblnShowByMonth = False
		maryOrderSummaryHeaders = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Year")
		pstrTableCaption = " for " & mstrYear
		ReDim maryOrderSummary(12)
		ReDim maryOrderSummary_PriorPeriod(12)
		
		For i = 0 To 11
			mdtFirstDayOfMonth = DateAdd("m", i, mdtFirstDayOfYear)
			mdtToday = DateAdd("d", -1, DateAdd("m", 1, mdtFirstDayOfMonth))
			maryOrderSummary(i) = GetOrderSummaries(mdtFirstDayOfMonth, mdtToday, False, cbytTopSellers)
			maryOrderSummary_PriorPeriod(i) = GetOrderSummaries(DateAdd("y", -1, mdtFirstDayOfMonth), DateAdd("m", -1, mdtToday), False, cbytTopSellers)
		Next 'i
		maryOrderSummary(12) = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)
		maryOrderSummary_PriorPeriod(12) = GetOrderSummaries(DateAdd("y", -1, mdtFirstDayOfYear), DateAdd("m", -1, mdtToday), False, cbytTopSellers)
	Case Else:
		pblnShowByMonth = False
		maryOrderSummaryHeaders = Array("Day", "Week", "Month", "Year")
		ReDim maryOrderSummary(3)
		maryOrderSummary(0) = GetOrderSummaries(mdtFirstDay, mdtToday, False, cbytTopSellers)
		maryOrderSummary(1) = GetOrderSummaries(mdtFirstDayOfWeek, mdtToday, False, cbytTopSellers)
		maryOrderSummary(2) = GetOrderSummaries(mdtFirstDayOfMonth, mdtToday, False, cbytTopSellers)
		maryOrderSummary(3) = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, cbytTopSellers)
End Select

pstrProductLink = "<li><prodName>&nbsp;<b>(<salesQty>)</b></li>"
If cblnAddon_ProductMgr Then pstrProductLink = "<li><a href='sfProductAdmin.asp?Action=ViewProduct&ViewID=<prodID>'><prodName></a>&nbsp;<b>(<salesQty>)</b></li>"
Response.Flush

If Not pblnShowByMonth Then
%>
<table class="tbl" border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" ID="tblOrderSummaries">
    <tr>
    <th rowspan="2">&nbsp;</th>
    <th colspan="<%= UBound(maryOrderSummaryHeaders) + 1 %>" class="hdrNonSelected">Order Summaries <%= pstrTableCaption %></th>
    </tr>
    <tr class="tdHighlight">
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1"><%= maryOrderSummaryHeaders(j) %></th>
    <% Next 'j %>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Orders</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryOrderSummary(j)(enOrderSummary_SalesCount) %>&nbsp;</td>
    <% Next 'j %>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Sales</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= FormatCurrency(maryOrderSummary(j)(enOrderSummary_SalesTotal), 2) %>&nbsp;</td>
    <% Next 'j %>
    </tr>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Avg. Order</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <td align="center" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= FormatCurrency(maryOrderSummary(j)(enOrderSummary_AverageOrder), 2) %>&nbsp;</td>
    <% Next 'j %>
    </tr>
    <% If cbytTopSellers > -1 Then %>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">Best Sellers</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <td align="left" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= writeTopSellers(maryOrderSummary(j)(enOrderSummary_TopSellers), pstrProductLink, cstrOpenTag, cstrCloseTag) %></td>
    <% Next 'j %>
    </tr>
    <% End If	'cbytTopSellers > -1 %>
</table>
<!--webbot bot="PurpleText" preview="End Order Reports" -->
<% End If	'pblnShowByMonth %>

<!--webbot bot="PurpleText" preview="Begin Product Sales Reports" -->
<% If pstrAction = "YearlyReport" Then %>
<br />
<table class="tbl" border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" ID="tblProductSalesSummaries">
<% If pblnShowByMonth Then %>
    <tr>
    <th colspan="2" style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;</th>
    <th colspan="<%= (UBound(maryOrderSummaryHeaders) + 1) * 24 + 4 %>" class="hdrNonSelected">Product Sales Summaries <%= pstrTableCaption %></th>
    </tr>
    <tr class="tdHighlight">
    <th rowspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Code</th>
    <th rowspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Name</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <th colspan="24" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1"><%= maryOrderSummaryHeaders(j) %></th>
    <% Next 'j %>
    <th rowspan="2" colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Totals</th>
    </tr>
    <tr class="tdHighlight">
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Jan</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Feb</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Mar</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Apr</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">May</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Jun</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Jul</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Aug</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Sep</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Oct</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Nov</th>
      <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Dec</th>
    <% Next 'j %>
    </tr>
<% Else %>
    <tr>
    <th colspan="2" style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;</th>
    <th colspan="<%= (UBound(maryOrderSummaryHeaders) + 1) * 2 + 2 %>" class="hdrNonSelected">Product Sales Summaries <%= pstrTableCaption %></th>
    </tr>
    <tr class="tdHighlight">
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Code</th>
    <th style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Name</th>
    <% For j = 0 To UBound(maryOrderSummaryHeaders) %>
    <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1"><%= maryOrderSummaryHeaders(j) %></th>
    <% Next 'j %>
    <th colspan="2" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">Totals</th>
    </tr>
<% End If %>

    <%
    ReDim paryColumnTotals(UBound(maryProducts(0)))
    For i = 0 To UBound(maryProducts)
		For j = 2 To UBound(maryProducts(i)) Step 2
			paryColumnTotals(j) = paryColumnTotals(j) + maryProducts(i)(j)
			paryColumnTotals(j + 1) = paryColumnTotals(j + 1) + maryProducts(i)(j + 1)
		Next 'j
	Next 'i
	
    For i = 0 To UBound(maryProducts)
    %>
    <tr>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryProducts(i)(0) %></th>
    <th align="left" valign="top" class="tdHighlight" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryProducts(i)(1) %></th>
    <% For j = 2 To UBound(maryProducts(i)) Step 2 %>
    <td align="center" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= maryProducts(i)(j) %></td>
    <td align="center" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= FormatCurrency(maryProducts(i)(j + 1), 2) %></td>
    <% Next 'j %>
    </tr>
    <% Next 'i %>
    <tr>
    <th colspan="2">&nbsp;</th>
    <% For j = 2 To UBound(paryColumnTotals) Step 2 %>
    <th align="center" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= paryColumnTotals(j) %></th>
    <th align="center" valign="top" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1"><%= FormatCurrency(paryColumnTotals(j + 1), 2) %></th>
    <% Next 'j %>
    </tr>

</table>
<% End If %>
<!--webbot bot="PurpleText" preview="End Product Sales Reports" -->

    </td>
  </tr>
  <tr>
    <td>
	  <!--webbot bot="PurpleText" preview="Begin Bottom Navigation Section" -->
      <!--#include file="adminFooter.asp"-->
	  <!--webbot bot="PurpleText" preview="End Bottom Navigation Section" -->
     </td>
  </tr>
</table>

<% End Sub	'Main %>