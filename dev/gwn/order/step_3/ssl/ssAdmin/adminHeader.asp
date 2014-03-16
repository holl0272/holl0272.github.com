<%
'********************************************************************************
'*   Common Support File			                                            *
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
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="ssLibrary/ssmodSF5Addons.asp"-->
<!--#include file="ssReports.asp"-->
<%
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
%>
<% Sub WriteHeaderOrderSummaries %>
<%
Const cblnShowArrows = True

Dim maryDailyOrderSummary
Dim maryWeeklyOrderSummary
Dim maryMonthlyOrderSummary
Dim maryYearlyOrderSummary
Dim maryDailyOrderSummary_PriorPeriod
Dim maryWeeklyOrderSummary_PriorPeriod
Dim maryMonthlyOrderSummary_PriorPeriod
Dim maryYearlyOrderSummary_PriorPeriod

'maryDailyOrderSummary = GetOrderSummaries(Date(), Date(), False)
'maryWeeklyOrderSummary = GetOrderSummaries(Date(), DateAdd("d", -7, Date()), False)

'Dim mlngStartDay:	mlngStartDay = vbSunday	'change to vbSunday to start on Sunday
Dim mlngStartDay:	mlngStartDay = vbMonday	'change to vbSunday to start on Sunday
Dim mdtToday:		mdtToday = Date()
Dim mbytToday
Dim mdtFirstDayOfWeek
Dim mdtFirstDayOfMonth
Dim mdtFirstDayOfYear

mbytToday = Weekday(mdtToday,mlngStartDay)
mdtFirstDayOfWeek = DateAdd("d", -1 * mbytToday + 1, mdtToday)
mdtFirstDayOfMonth = Month(mdtToday) & "/1/" & Year(mdtToday)
mdtFirstDayOfYear = "1/1/" & Year(mdtToday)

If False Then
	Response.Write "<fieldset><legend>dates</legend>"
	Response.Write "Today: " & mdtToday & "<br />"
	Response.Write "Day: " & mbytToday & "<br />"
	Response.Write "mlngStartDay: " & mlngStartDay & "<br />"
	Response.Write "mdtFirstDayOfWeek: " & mdtFirstDayOfWeek & "<br />"
	Response.Write "mdtFirstDayOfMonth: " & mdtFirstDayOfMonth & "<br />"
	Response.Write "mdtFirstDayOfYear: " & mdtFirstDayOfYear & "<br />"
	Response.Write "</fieldset>"
End If

maryDailyOrderSummary = GetOrderSummaries(mdtToday, mdtToday, False, -1)
maryWeeklyOrderSummary = GetOrderSummaries(mdtFirstDayOfWeek, mdtToday, False, -1)
maryMonthlyOrderSummary = GetOrderSummaries(mdtFirstDayOfMonth, mdtToday, False, -1)
maryYearlyOrderSummary = GetOrderSummaries(mdtFirstDayOfYear, mdtToday, False, -1)

maryDailyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("d", -1, mdtToday), DateAdd("d", -1, mdtToday), False, -1)
maryWeeklyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("d", -7, mdtFirstDayOfWeek), DateAdd("d", -7, mdtToday), False, -1)
maryMonthlyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("m", -1, mdtFirstDayOfMonth), DateAdd("m", -1, mdtToday), False, -1)
maryYearlyOrderSummary_PriorPeriod = GetOrderSummaries(DateAdd("y", -1, mdtFirstDayOfYear), DateAdd("y", -1, mdtToday), False, -1)

%>    
<table border="0" cellspacing="0" cellpadding="2" style="border-collapse:collapse;font-size:8pt;display:inline;text-align:right">
<tbody>
<tr class="tdHighlight">
  <th style="border-bottom:solid 1pt black;background:white;">&nbsp;</th>
  <th style="border-top:solid 1pt black;border-bottom:solid 1pt black;border-left:solid 1pt black;text-align:center">Day</th>
  <th style="border-top:solid 1pt black;border-right:solid 1pt black;border-left:solid 1pt black;border-bottom:solid 1pt black;text-align:center">Week</th>
  <th style="border-top:solid 1pt black;border-right:solid 1pt black;border-left:solid 1pt black;border-bottom:solid 1pt black;text-align:center">Month</th>
  <th style="border-top:solid 1pt black;border-right:solid 1pt black;border-left:solid 1pt black;border-bottom:solid 1pt black;text-align:center">Year</th>
</tr>
<tr>
  <td class="tdHighlight" style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black">Sales:</td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_SalesTotal), maryDailyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_SalesTotal), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_SalesTotal), maryMonthlyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_SalesTotal), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_SalesTotal), True, cblnShowArrows) %></td>
</tr>
<tr>
  <td class="tdHighlight" style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black">Orders: </td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_SalesCount), maryDailyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_SalesCount), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_SalesCount), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_SalesCount), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_SalesCount), False, cblnShowArrows) %></td>
</tr>
<tr>
  <td class="tdHighlight" style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black">Avg: </td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryDailyOrderSummary(enOrderSummary_AverageOrder), maryDailyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryWeeklyOrderSummary(enOrderSummary_AverageOrder), maryWeeklyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryMonthlyOrderSummary(enOrderSummary_AverageOrder), maryMonthlyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), True, cblnShowArrows) %></td>
  <td style="border-left:solid 1pt black;border-bottom:solid 1pt black;border-right:solid 1pt black"><%= WriteDifference(maryYearlyOrderSummary(enOrderSummary_AverageOrder), maryYearlyOrderSummary_PriorPeriod(enOrderSummary_AverageOrder), True, cblnShowArrows) %></td>
</tr></tbody>
</table>

<% End Sub	'WriteHeaderOrderSummaries %>

<% Sub WriteHeader(strBodyOnload, blnDisplayTopMenu) %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<%
	On Error Resume Next
	Response.Expires = 60
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	If Err.number <> 0 Then Err.Clear
%>
<title><%= mstrPageTitle %></title>
<link rel="stylesheet" href="ssLibrary/ssStyleSheet.css" type="text/css">
<link rel="stylesheet" href="dtree/dtree.css" type="text/css">
<script language="javascript" src="dtree/dtree.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/ssFormValidation.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/calendar.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/sorter.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/tipMessage.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/ssm.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

function showDataEntryTip(theLabel)
{
	stm(tipMessage[theLabel.htmlFor],tipStyle['dataEntry']);
}

	var FiltersEnabled = 0 // if your not going to use transitions or filters in any of the tips set this to 0
	tipStyle['dataEntry']=["white","black","steelblue","whitesmoke","","","","","","","","","","",200,"",2,2,10,10,"","","","simple","gray"]

	//original examples
	// tipStyle[...]=[TitleColor,TextColor,TitleBgColor,TextBgColor,TitleBgImag,TextBgImag,TitleTextAlign,TextTextAlign, TitleFontFace, TextFontFace, TipPosition, StickyStyle, TitleFontSize, TextFontSize, Width, Height, BorderSize, PadTextArea, CoordinateX , CoordinateY, TransitionNumber, TransitionDuration, TransparencyLevel ,ShadowType, ShadowColor]
	tipStyle[0]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,1,0,"",""]
	tipStyle[1]=["white","black","#000099","#E8E8FF","","","","","","","center","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[2]=["white","black","#000099","#E8E8FF","","","","","","","left","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[3]=["white","black","#000099","#E8E8FF","","","","","","","float","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[4]=["white","black","#000099","#E8E8FF","","","","","","","fixed","","","",200,"",2,2,1,1,"","","","",""]
	tipStyle[5]=["white","black","#000099","#E8E8FF","","","","","","","","sticky","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[6]=["white","black","#000099","#E8E8FF","","","","","","","","keep","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[7]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,40,10,"","","","",""]
	tipStyle[8]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,50,"","","","",""]
	tipStyle[9]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,75,"simple","gray"]
	tipStyle[10]=["white","black","black","white","","","right","","Impact","cursive","center","",3,5,200,150,5,20,10,0,50,1,80,"complex","gray"]
	tipStyle[11]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,45,"simple","gray"]
	tipStyle[12]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,"","","","",""]

</script>
</head>
<% 
If Len(strBodyOnload) > 0 Then
	Response.Write "<body onload=" & Chr(34) & strBodyOnload & Chr(34) & ">"
Else
	Response.Write "<body>"
End If
%>
<div id="TipLayer" style="visibility:hidden;position:absolute;z-index:1000;top:-100"></div>
<%
If blnDisplayTopMenu Then
	'Call WriteTreeMenu
%>
<div id="topMenu">
	<map name="FPMap0" id="FPMap0">
	<area href="http://www.gamewearnow.com" shape="rect" coords="0, 0, 78, 81">
	</map>
	<img alt="GWN Circle" src="GWN_121406.gif" width="82" height="82" usemap="#FPMap0" border="0px">
	<% If isAllowedToViewReports Then %><a class="topMenu" href="default.asp">Dashboard</a><% End If	'isAllowedToViewReports %>
	<a class="topMenu" href="admin.asp">Main&nbsp;Menu</a>
	<% If isAllowedToViewOrder Then %><a class="topMenu" href="ssOrderAdmin.asp">Orders</a><% End If	'isAllowedToViewOrder %>
	<% If isAllowedToEditProducts Then %><a class="topMenu" href="sfProductAdmin.asp">Catalog</a><% End If	'isAllowedToEditProducts %>
	<a class="topMenu" href="ssHelpFiles/help.htm">Help</a>
	<% If isLoggedIn Then Response.Write " <a class=""topMenu"" href=""Admin.asp?Action=LogOff"">Log Off " & userName & "</a>" %>
	<% 'If isLoggedIn Then Call WriteHeaderOrderSummaries %>
</div>
<% End If	'blnDisplayTopMenu %>
<!-- End Sandshot Header -->
<% End Sub	'WriteHeader 

'*******************************************************************************************************************************************

Dim maryTreeMenuItems

Sub createTreeMenu

Dim i
Dim menuIndex
Dim parentIndex

	If cblnSF5AE Then
		ReDim maryTreeMenuItems(70)
	Else
		ReDim maryTreeMenuItems(68)
	End If
	menuIndex = 0
	
	maryTreeMenuItems(menuIndex) = Array(-1, Application("adminStoreName") & " Store Menu&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", ""):	menuIndex = menuIndex + 1
	maryTreeMenuItems(menuIndex) = Array(0, "Orders", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "View Orders", "ssOrderAdmin.asp"):	menuIndex = menuIndex + 1
		'maryTreeMenuItems(menuIndex) = Array(parentIndex, "Process Orders", "ssOrderAdmin_Process.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Mass Process Payments", "ssOrderAdmin.asp?Action=ViewOrder&optDisplay=1&optDate_Filter=0&optPayment_Filter=2&optShipment_Filter=0"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Mass Process Shipments", "ssOrderAdmin.asp?Action=ViewOrder&optDisplay=2&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=1"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Import Tracking Numbers", "ssOrderAdmin.asp?Action=Filter&optDisplay=3"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Update Old Orders", "ssOrderAdmin_UpdatePastOrders.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Customers", "sfCustomerAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Affiliates", "sfAffiliatesAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Discounts and Promotions", "ssPromotionsAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Buyer Club Remptions", "ssBuyersClubRedemptionAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Gift Certificates", "ssGiftCertificateAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "by Order: <form name=""frmQuickOrderManager"" id=""frmQuickOrderManager"" action=""ssOrderAdmin.asp"" style=""display: inline""><input type=""hidden"" name=""Action"" id=""Action"" value=""Filter""><input type=""hidden"" name=""Flag_Voided"" id=""Flag_Voided"" value=""0""><input type=""hidden"" name=""optText_Filter"" id=""optText_Filter1"" value=""1""><input type=""text"" class=""dtree"" name=""Text_Filter"" id=""Text_Filter1"" size=""10"" value="""">&nbsp;<input type=""image"" src=""../../Images/Buttons/go3.gif""></form>", ""):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "by Email: <form name=""frmQuickOrderManager"" id=""frmQuickOrderManager"" action=""ssOrderAdmin.asp"" style=""display: inline""><input type=""hidden"" name=""Action"" id=""Action"" value=""Filter""><input type=""hidden"" name=""Flag_Voided"" id=""Flag_Voided"" value=""0""><input type=""hidden"" name=""optText_Filter"" id=""optText_Filter1"" value=""4""><input type=""text"" class=""dtree"" name=""Text_Filter"" id=""Text_Filter4"" size=""10"" value="""">&nbsp;<input type=""image"" src=""../../Images/Buttons/go3.gif""></form>", ""):	menuIndex = menuIndex + 1

	maryTreeMenuItems(menuIndex) = Array(0, "Products", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Product Administration", "sfProductAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Product Placement", "ssProductPlacementAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Product Export", "ssProductExportTool.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Product Pricing Tool", "ssProductPricingTool.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Import Products", "ssImportProducts.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Import Data", "ssImportUtility.asp"):	menuIndex = menuIndex + 1
		If cblnSF5AE Then maryTreeMenuItems(menuIndex) = Array(parentIndex, "Load Inventory File", "ssInventoryImportFile.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Check Images", "ssProductImageCheck.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Product Reviews", "ssProductReviewsAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Pricing Levels", "ssPricingLevelAdmin.asp"):	menuIndex = menuIndex + 1
		If cblnSF5AE Then
			maryTreeMenuItems(menuIndex) = Array(parentIndex, "Categories", "sfCategoryAdminAE.asp"):	menuIndex = menuIndex + 1
		Else
			maryTreeMenuItems(menuIndex) = Array(parentIndex, "Categories", "sfCategoryAdmin.asp"):	menuIndex = menuIndex + 1
		End If
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Manufacturers", "sfManufacturersAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Vendors", "sfVendorsAdmin.asp"):	menuIndex = menuIndex + 1

	maryTreeMenuItems(menuIndex) = Array(0, "Reports", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Sales Central", "default.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Sales Report", "ssSalesReports.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Top Products", "ssSalesReports.asp?chkShowReport0=1"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Top Viewed Products", "ssSalesReports.asp?chkShowReport8=1"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Top Customers", "ssSalesReports.asp?chkShowReport7=1"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Sales Report, Advanced", "ssSalesReportAdvanced.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Annual Reports", "ssAnnualSalesReports.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Standard SF Reporting Tools", "../admin/sfreports.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Promotional Mail Manager", "ssPromoMailAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Monitor Site", "ssSiteMonitor.htm"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "PayPal Payments", "ssPayPalPaymentsAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Active Carts", "ssActiveShoppingCarts.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Saved Carts", "ssActiveWishLists.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Notification Requests", "ssActiveNotifications.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Discounts Usage", "ssPromotionsReport.asp"):	menuIndex = menuIndex + 1
		If cblnSF5AE Then maryTreeMenuItems(menuIndex) = Array(parentIndex, "List Inventory", "ssInventoryList.asp"):	menuIndex = menuIndex + 1

	maryTreeMenuItems(menuIndex) = Array(0, "Content Administration", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Page Content", "ssCMS_PageFragmentAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Manufacturers", "ssCMS_ManufacturerAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Content Types", "ssCMS_ContentTypeAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Generic Content Administration", "ssCMS_GenericContentAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Design Administration", "ssCMS_UserType1Admin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Theme Administration", "ssCMS_UserType2Admin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Breed Administration", "ssCMS_UserType3Admin.asp"):	menuIndex = menuIndex + 1

	maryTreeMenuItems(menuIndex) = Array(0, "Adminstrative Settings", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Change Username/Password", "Admin.asp?Action=ChangePwd"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Users", "ssUserAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Review Login Attempts", "ssUserLoginAttemptsAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Clean Up/Back-up Database", "ssDBcleanup.asp"):	menuIndex = menuIndex + 1

	maryTreeMenuItems(menuIndex) = Array(0, "Site Settings", ""):	menuIndex = menuIndex + 1
		parentIndex = menuIndex - 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Store Configuration Settings", "sfAdmin.asp?Show=Application"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Configure Colors", "sfColorAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Configure Fonts", "sfFontAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Design Settings", "sfDesignAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Geographical Settings", "sfAdmin.asp?Show=Geographical"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Mail Settings", "sfAdmin.asp?Show=Mail"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Search Result Text Settings", "sfTextAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Shipping Settings", "sfAdmin.asp?Show=Shipping"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Countries, Tax Rate", "sfLocalesCountryAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "States/Provinces, Tax Rate", "sfLocalesStateAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Local Tax Settings", "ssTaxRateAdmin.asp"):	menuIndex = menuIndex + 1
		maryTreeMenuItems(menuIndex) = Array(parentIndex, "Transaction Settings", "sfAdmin.asp?Show=Transaction"):	menuIndex = menuIndex + 1

	If UBound(maryTreeMenuItems) > menuIndex - 1 Then
		Response.Write "alert('Please redimension the menu array. It should only have " & menuIndex - 1 & " items.');" & vbcrlf
	End If

	For i = 0 To menuIndex - 1
		Response.Write "d.add(" & i & "," & maryTreeMenuItems(i)(0) & ",'" & maryTreeMenuItems(i)(1) & "','" & maryTreeMenuItems(i)(2) & "');" & vbcrlf	
	Next 'i

End Sub	'createTreeMenu

'*******************************************************************************************************************************************

Sub WriteTreeMenu
%>
	<script type="text/javascript">
		<!--
		function dtreeIndex()
		{
			index++;
			return index;
		}
		
		var index = -1;
		
		d = new dTree('d');
		<% createTreeMenu %>

		//document.write('<div id="myMenu" style="border:solid 1 black">' + d + '</div>');

		ssmItems[0]=["Menu"] //create header
		buildMenu_Custom()
		function buildMenu_Custom()
		{
			if (IE||NS6)
			{
				document.write('<div id="basessm" style="visibility:hidden;Left : '+XOffset+'px ;Top : '+YOffset+'px ;width:'+(menuWidth+barWidth+10)+'px">')
				document.write('<div id="thessm" style="Left : '+(-menuWidth)+'px ;Top : 0 ;" onmouseover="moveOut()" onmouseout="moveBack()">')
				document.write('<table id="tblsm" border="0" cellpadding="0" cellspacing="0">');
				document.write('<tr><td>')
				document.write('<div id="divOuterMenu">' + d + '</div>');
				document.write('</td><td class="tdsm">M e n u</td></tr>')
				document.write('</table>')
				document.write('</div></div>')
			}
			theleft=-menuWidth;lastY=0;setTimeout('initSlide();', 1)
		}

		//-->
	</script>
<%
End Sub	'WriteTreeMenu
'*******************************************************************************************************************************************

'Determine which add-ons are installed
'Known possibilities include:

'Dim cblnAddon_ImportProducts	'Import Products
Dim cblnAddon_PayPalPayments	'PayPal Payments
Dim cblnAddon_PostageRate		'Postage Rate Component
Dim cblnAddon_PricingLevelMgr	'Pricing Level Manager
Dim cblnAddon_ProductMgr		'Product Manager
'Dim cblnAddon_ProductPricing	'Product Pricing
Dim cblnAddon_PromoMailMgr		'Promotional Mail Manager
Dim cblnAddon_PromoMgr			'Promotion Manager
Dim cblnAddon_PromotionMgrII	'Promotion Manager
'Dim cblnAddon_SalesCentral		'Sales Central
Dim cblnAddon_SiteMonitor		'Site Monitor
Dim cblnAddon_TaxRateMgr		'Tax Rate Manager
Dim cblnAddon_WebStoreMgr		'WebStore Manager
Dim cblnAddon_ZBS				'Zone Based Shipping
Dim cblnAddon_ProductPlacement	'Product Placement
Dim cblnAddon_ProductReview		'Product Review

Call DetermineAddOns

Sub DetermineAddOns

Dim pobjFSO
Dim pstrFilePath

	'On Error Resume Next

	pstrFilePath = ssAdminPath

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	
	'cblnAddon_ImportProducts = pobjFSO.FileExists(pstrFilePath & "ssImportProducts.asp")
	cblnAddon_PayPalPayments = pobjFSO.FileExists(pstrFilePath & "ssPayPalPaymentsAdmin.asp")
	cblnAddon_PostageRate = pobjFSO.FileExists(pstrFilePath & "ssPostageRate_shippingMethodsAdmin.asp")
	cblnAddon_PricingLevelMgr = pobjFSO.FileExists(pstrFilePath & "ssPricingLevelAdmin.asp")
	'cblnAddon_ProductExport = pobjFSO.FileExists(pstrFilePath & "ssProductExportTool.asp")
	cblnAddon_ProductMgr = pobjFSO.FileExists(pstrFilePath & "sfProductAdmin.asp")
	'cblnAddon_ProductPricing = pobjFSO.FileExists(pstrFilePath & "ssProductPricingTool.asp")
	cblnAddon_PromoMailMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoMailAdmin.asp")
	cblnAddon_PromoMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoAdmin.asp")
	cblnAddon_PromotionMgrII = pobjFSO.FileExists(pstrFilePath & "ssPromotionsAdmin.asp")
	If cblnAddon_PromotionMgrII Then cblnAddon_PromoMgr = False
	'cblnAddon_SalesCentral = pobjFSO.FileExists(pstrFilePath & "ssReports.asp")
	cblnAddon_SiteMonitor = pobjFSO.FileExists(pstrFilePath & "ssSiteMonitor.htm")
	cblnAddon_TaxRateMgr = pobjFSO.FileExists(pstrFilePath & "ssTaxRateAdmin.asp")
	cblnAddon_WebStoreMgr = pobjFSO.FileExists(pstrFilePath & "sfDesignAdmin.asp")
	cblnAddon_ZBS = pobjFSO.FileExists(pstrFilePath & "sszbsZoneAdmin.asp")
	cblnAddon_ProductPlacement = pobjFSO.FileExists(pstrFilePath & "ssProductPlacementAdmin.asp")
	cblnAddon_ProductReview = pobjFSO.FileExists(pstrFilePath & "ssProductReviewsAdmin.asp")
	'debugprint "cblnAddon_ProductPlacement",cblnAddon_ProductPlacement

	Set pobjFSO = Nothing

End Sub	'DetermineAddOns
'
%>