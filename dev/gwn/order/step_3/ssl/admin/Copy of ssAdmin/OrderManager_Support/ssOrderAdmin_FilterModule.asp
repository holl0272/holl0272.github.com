<%
'********************************************************************************
'*   Order Manager							                                    *
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
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'Filter Elements
Dim mbytText_Filter
Dim mstrText_Filter

Dim mbytDate_Filter
Dim mbytDateType_Filter
Dim mstrStartDate
Dim mstrEndDate

Dim mbytPayment_Filter
Dim mbytShipment_Filter
Dim mbytoptFlag_Filter
Dim mbytoptBackOrdered
Dim mbytFlag_ssExported
Dim mbytFlag_ssExportedPayment
Dim mbytFlag_ssssExportedShipping
Dim mbytFlag_orderVoided
Dim mbytIncomplete_Filter
Dim mbytOrderStatus_Filter
Dim mbytInternalOrderStatus_Filter
Dim mstrPaymentType_Filter
Dim mbytQuick_Filter
Dim mstrMfg_Filter
Dim mstrVend_Filter

Dim mstrProductID

Const cblnUseCustomFilter = False
Dim mbytCustomFilter

'This section defines which fields appear for the filter by text
Dim maryTextFilterOptions(13)
maryTextFilterOptions(0) = "Do Not Include"
maryTextFilterOptions(1) = "Order #"
maryTextFilterOptions(2) = "Last Name - Billing"
maryTextFilterOptions(3) = "First Name - Billing"
maryTextFilterOptions(4) = "Email - Billing"
maryTextFilterOptions(5) = "Company - Billing"
maryTextFilterOptions(6) = "Address - Billing"
maryTextFilterOptions(7) = "Last Name - Shipping"
maryTextFilterOptions(8) = "First Name - Shipping"
maryTextFilterOptions(9) = "Email - Shipping"
maryTextFilterOptions(10) = "Company - Shipping"
maryTextFilterOptions(11) = "Address - Shipping"
maryTextFilterOptions(12) = "Affiliate"
maryTextFilterOptions(13) = "Special Instructions"

'This section defines which fields appear for the filter by text
Dim maryDateType_Filter(2)
maryDateType_Filter(0) = Array("Order Date", "orderDate")
maryDateType_Filter(1) = Array("Paid Date", "ssDatePaymentReceived")
maryDateType_Filter(2) = Array("Shipping Date", "ssDateOrderShipped")

'***********************************************************************************************

Function SummaryFilter

Dim paryTemp
Dim pstrsqlWhere
Dim pstrsqlGroupBy
Dim pstrsqlHaving
Dim pstrOrderBy
Dim pstrTemp

	'load the incomplete filter
	mbytIncomplete_Filter = LoadRequestValue("optIncomplete_Filter")
	If len(mbytIncomplete_Filter) = 0 Then mbytIncomplete_Filter = 0
	Select Case mbytIncomplete_Filter
		Case "0"	'Complete
			pstrsqlWhere = " Where (sfOrders.orderIsComplete=1)"
			pstrsqlHaving = " Having (sfOrders.orderIsComplete=1)"
			pstrsqlGroupBy = ""
		Case "1"	'All
			pstrsqlWhere = " Where (((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null)))"
			pstrsqlHaving = " Having (((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null)))"
			pstrsqlGroupBy = ""
		Case "2"	'Incomplete
			pstrsqlWhere = " Where ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete is Null))"
			pstrsqlHaving = " Having ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete is Null))"
			pstrsqlGroupBy = ""
	End Select	

	'load the text filter
	mbytText_Filter = LoadRequestValue("optText_Filter")
	mstrText_Filter = LoadRequestValue("Text_Filter")

	'hack to save direct link in capability: Action=ViewOrder&OrderID=xxx
	If CBool(Request.QueryString("Action") = "ViewOrder") And CBool(Len(Request.QueryString("OrderID")) > 0) Then
		mbytText_Filter = "1"
		mstrText_Filter = Request.QueryString("OrderID")
	End If

	If len(mstrText_Filter) > 0 Then
		Select Case mbytText_Filter
			Case "0"	'Do Not Include
			Case "1"	'Order
				If InStr(1, mstrText_Filter, ",") > 0 Then
					paryTemp = Split(mstrText_Filter, ",")
					pstrTemp = "(orderID Like '%" & Trim(sqlSafe(paryTemp(0))) & "%')"
					For i = 1 To UBound(paryTemp)
						pstrTemp = pstrTemp & " Or (orderID Like '%" & Trim(sqlSafe(paryTemp(i))) & "%')"
					Next 'i
					pstrsqlWhere = pstrsqlWhere & " AND (" & pstrTemp & ")"
				ElseIf InStr(1, mstrText_Filter, ";") > 0 Then
					paryTemp = Split(mstrText_Filter, ";")
					pstrTemp = "(orderID Like '%" & Trim(sqlSafe(paryTemp(0))) & "%')"
					For i = 1 To UBound(paryTemp)
						pstrTemp = pstrTemp & " Or (orderID Like '%" & Trim(sqlSafe(paryTemp(i))) & "%')"
					Next 'i
					pstrsqlWhere = pstrsqlWhere & " AND (" & pstrTemp & ")"
					pstrsqlWhere = pstrsqlWhere & " AND (orderID Like '%" & sqlSafe(mstrText_Filter) & "%')"
				Else
					pstrsqlWhere = pstrsqlWhere & " AND (orderID Like '%" & sqlSafe(mstrText_Filter) & "%')"
				End If
			Case "2"	'Last Name
				pstrsqlWhere = pstrsqlWhere & " AND (custLastName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "3"	'First Name
				pstrsqlWhere = pstrsqlWhere & " AND (custFirstName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "4"	'email
				pstrsqlWhere = pstrsqlWhere & " AND (custEmail Like '%" & mstrText_Filter & "%')"
			Case "5"	'Company Name
				pstrsqlWhere = pstrsqlWhere & " AND (custCompany Like '%" & mstrText_Filter & "%')"
			Case "6"	'Address
				pstrsqlWhere = pstrsqlWhere & " AND ((custAddr1 Like '%" & mstrText_Filter & "%')) Or ((custAddr2 Like '%" & mstrText_Filter & "%'))"
			Case "7"	'Last Name
				pstrsqlWhere = pstrsqlWhere & " AND (cshpaddrShipLastName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "8"	'First Name
				pstrsqlWhere = pstrsqlWhere & " AND (cshpaddrShipFirstName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "9"	'email
				pstrsqlWhere = pstrsqlWhere & " AND (cshpaddrShipEmail Like '%" & mstrText_Filter & "%')"
			Case "10"	'Company Name
				pstrsqlWhere = pstrsqlWhere & " AND (cshpaddrShipCompany Like '%" & mstrText_Filter & "%')"
			Case "11"	'Address
				pstrsqlWhere = pstrsqlWhere & " AND ((cshpaddrShipAddr1 Like '%" & mstrText_Filter & "%')) Or ((cshpaddrShipAddr2 Like '%" & mstrText_Filter & "%'))"
			Case "12"	'Affilliate
				pstrsqlWhere = pstrsqlWhere & " AND (orderTradingPartner Like '%" & mstrText_Filter & "%')"
			Case "13"	'Special Instructions
				pstrsqlWhere = pstrsqlWhere & " AND (orderComments Like '%" & mstrText_Filter & "%')"
		End Select	
	End If

	'load the radio filters
	mbytPayment_Filter = LoadRequestValue("optPayment_Filter")
	mbytShipment_Filter = LoadRequestValue("optShipment_Filter")
	mbytDate_Filter = LoadRequestValue("optDate_Filter")
	mbytDateType_Filter = LoadRequestValue("optDateType_Filter")
	mbytoptFlag_Filter	= LoadRequestValue("optFlag_Filter")
	mbytoptBackOrdered	= LoadRequestValue("optBackOrdered")
	mbytFlag_ssExported = LoadRequestValue("Flag_ssExported")
	mbytFlag_ssExportedPayment = LoadRequestValue("Flag_ssExportedPayment")
	mbytFlag_ssssExportedShipping = LoadRequestValue("Flag_ssssExportedShipping")
	mbytFlag_orderVoided = LoadRequestValue("Flag_orderVoided")
	mbytOrderStatus_Filter = LoadRequestValue("Flag_OrderStatus")
	mbytInternalOrderStatus_Filter = LoadRequestValue("Flag_InternalOrderStatus")
	mstrPaymentType_Filter = LoadRequestValue("optPaymentType")
	mbytQuick_Filter = LoadRequestValue("optQuick_Filter")

	'set the defaults
	If len(mbytPayment_Filter) = 0 Then mbytPayment_Filter = 0
	If len(mbytShipment_Filter) = 0 Then mbytShipment_Filter = 0
	If len(mbytDate_Filter) = 0 Then mbytDate_Filter = 0
	If len(mbytDateType_Filter) = 0 Then mbytDateType_Filter = 0
	If len(mbytoptFlag_Filter) = 0 Then mbytoptFlag_Filter = 0
	If len(mbytoptBackOrdered) = 0 Then mbytoptBackOrdered = 0
	If len(mbytFlag_ssExported) = 0 Then mbytFlag_ssExported = 0
	If len(mbytFlag_ssExportedPayment) = 0 Then mbytFlag_ssExportedPayment = 0
	If len(mbytFlag_ssssExportedShipping) = 0 Then mbytFlag_ssssExportedShipping = 0
	If len(mbytFlag_orderVoided) = 0 Then mbytFlag_orderVoided = 0
	If len(mbytOrderStatus_Filter) = 0 Then mbytOrderStatus_Filter = -1
	If len(mbytInternalOrderStatus_Filter) = 0 Then mbytInternalOrderStatus_Filter = -1
	If len(mbytQuick_Filter) = 0 Then mbytQuick_Filter = 0
	
	If Len(Request.Form & Request.QueryString) = 0 Then
		mbytDate_Filter = 1
		mstrStartDate = DateAdd("d",-1,Date())
	Else
		mstrStartDate = LoadRequestValue("StartDate")
	End If
	If Len(mstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (" & maryDateType_Filter(mbytDateType_Filter)(1) & " >= " & wrapSQLValue(mstrStartDate & " 12:00:00 AM", True, enDatatype_date) & ")"
	
	mstrEndDate = LoadRequestValue("EndDate")
	If Len(mstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (" & maryDateType_Filter(mbytDateType_Filter)(1) & " <= " & wrapSQLValue(mstrEndDate & " 11:59:59 PM", True, enDatatype_date) & ")"

	If CLng(mbytOrderStatus_Filter) <> -1 Then
		Select Case CLng(mbytOrderStatus_Filter)
			Case 0	'Unread/default
				pstrsqlWhere = pstrsqlWhere & " and (ssOrderStatus=" & mbytOrderStatus_Filter & " OR ssOrderStatus is Null)"
			Case Else
				pstrsqlWhere = pstrsqlWhere & " and (ssOrderStatus=" & mbytOrderStatus_Filter & ")"
		End Select	
	End If	'CLng(mbytOrderStatus_Filter) <> -1

	If CLng(mbytInternalOrderStatus_Filter) <> -1 Then
		Select Case CLng(mbytInternalOrderStatus_Filter)
			Case 0	'Unread/default
				pstrsqlWhere = pstrsqlWhere & " and (ssInternalOrderStatus=" & mbytInternalOrderStatus_Filter & " OR ssOrderStatus is Null)"
			Case Else
				pstrsqlWhere = pstrsqlWhere & " and (ssInternalOrderStatus=" & mbytInternalOrderStatus_Filter & ")"
		End Select	
	End If	'CLng(mbytInternalOrderStatus_Filter) <> -1

	If Len(mstrPaymentType_Filter) > 0 Then
		If mstrPaymentType_Filter = "zilch" Then
			pstrsqlWhere = pstrsqlWhere & " and ((orderPaymentMethod is null) OR (orderPaymentMethod=''))"
		Else
			pstrsqlWhere = pstrsqlWhere & " and (orderPaymentMethod=" & wrapSQLValue(mstrPaymentType_Filter, True, enDatatype_string) & ")"
		End If
	End If

	Select Case mbytPayment_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " and ssDatePaymentReceived is Not Null"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " and ssDatePaymentReceived is Null"
	End Select	

	Select Case mbytoptFlag_Filter
		Case "0"	'Do Not Include
		Case "1"	'Flagged
			pstrsqlWhere = pstrsqlWhere & " and ssOrderFlagged=1"
		Case "2"	'unflagged
			pstrsqlWhere = pstrsqlWhere & " and (ssOrderFlagged=0 OR ssOrderFlagged is Null)"
	End Select	

	Select Case mbytoptBackOrdered
		Case "0"	'Do Not Include
		Case "1"	'Back Ordered  
			pstrsqlWhere = pstrsqlWhere & " and ssBackOrderDateExpected is Not Null"
		Case "2"	'Not BackOrdered 
			pstrsqlWhere = pstrsqlWhere & " and (ssBackOrderDateExpected=0 OR ssBackOrderDateExpected is Null)"
	End Select	

	Select Case mbytFlag_ssExported
		Case "0"	'Do Not Include
		Case "1"	'Back Ordered  
			pstrsqlWhere = pstrsqlWhere & " and ssExported is Not Null"
		Case "2"	'Not BackOrdered 
			pstrsqlWhere = pstrsqlWhere & " and (ssExported=0 OR ssExported is Null)"
	End Select	

	Select Case mbytFlag_ssExportedPayment
		Case "0"	'Do Not Include
		Case "1"	'Back Ordered  
			pstrsqlWhere = pstrsqlWhere & " and ssExportedPayment is Not Null"
		Case "2"	'Not BackOrdered 
			pstrsqlWhere = pstrsqlWhere & " and (ssExportedPayment=0 OR ssExportedPayment is Null)"
	End Select	

	Select Case mbytFlag_ssssExportedShipping
		Case "0"	'Do Not Include
		Case "1"	'Back Ordered  
			pstrsqlWhere = pstrsqlWhere & " and ssssExportedShipping is Not Null"
		Case "2"	'Not BackOrdered 
			pstrsqlWhere = pstrsqlWhere & " and (ssssExportedShipping=0 OR ssssExportedShipping is Null)"
	End Select	

	Select Case mbytFlag_orderVoided
		Case "0"	'Do Not Include
		Case "1"	'Voided  
			pstrsqlWhere = pstrsqlWhere & " and sfOrders.orderVoided"
		Case "2"	'Not Voided
			pstrsqlWhere = pstrsqlWhere & " and (sfOrders.orderVoided=0 OR sfOrders.orderVoided is Null)"
	End Select	

	'Custom Filter
	mbytCustomFilter = LoadRequestValue("optCustomFilter")
		If len(mbytCustomFilter) = 0 Then mbytCustomFilter = "0"
	
	If True Then
		Select Case mbytCustomFilter
			Case "0"	'Do Not Include
			Case "1"	'default store
				pstrsqlWhere = pstrsqlWhere & " and ((orderStoreID='SMD') or (orderStoreID is Null))"
				pstrsqlGroupBy = pstrsqlGroupBy & ", sfOrders.orderStoreID"
			Case "2"
				pstrsqlWhere = pstrsqlWhere & " and orderStoreID='meu'"
				pstrsqlGroupBy = pstrsqlGroupBy & ", sfOrders.orderStoreID"
		End Select	
	Else
		Select Case mbytCustomFilter
			Case "0"	'Do Not Include
			Case "1"	'default store
				pstrsqlWhere = pstrsqlWhere & " and ((storeID=1) or (storeID is Null))"
				pstrsqlGroupBy = pstrsqlGroupBy & ", sfOrders.storeID"
			Case Else
				pstrsqlWhere = pstrsqlWhere & " and storeID=" & mbytCustomFilter
				pstrsqlGroupBy = pstrsqlGroupBy & ", sfOrders.storeID"
		End Select	
	End If
	
	Select Case mbytShipment_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " and ssDateOrderShipped is Null"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " and ssDateOrderShipped is Not Null"
	End Select	
	
	'load the Product ID filter
	mstrProductID = LoadRequestValue("ProductID")
	If Len(mstrProductID) > 0 Then
		paryTemp = Split(mstrProductID,", ")
		pstrsqlWhere = pstrsqlWhere & " And (sfOrderDetails.odrdtProductID=" & wrapSQLValue(paryTemp(0), True, enDatatype_string)
		For i = 1 To UBound(paryTemp)
			pstrsqlWhere = pstrsqlWhere & " OR sfOrderDetails.odrdtProductID=" & wrapSQLValue(paryTemp(1), True, enDatatype_string)
		Next 'i
		pstrsqlWhere = pstrsqlWhere & ")"
		'pstrJoinType = "INNER"
	End If
		
	'load the mfg filter
	mstrMfg_Filter = LoadRequestValue("Mfg_Filter")
	If Len(mstrMfg_Filter) > 0 Then
		pstrsqlWhere = pstrsqlWhere & " And (sfOrderDetails.odrdtManufacturer=" & wrapSQLValue(mstrMfg_Filter, True, enDatatype_string) & ")"
	End If
		
	'load the vend filter
	mstrVend_Filter = LoadRequestValue("Vend_Filter")
	If Len(mstrVend_Filter) > 0 Then
		pstrsqlWhere = pstrsqlWhere & " And (sfOrderDetails.odrdtVendor=" & wrapSQLValue(mstrVend_Filter, True, enDatatype_string) & ")"
	End If
		
	'Build  the Order By
	mstrOrderBy = LoadRequestValue("OrderBy")
	If len(mstrOrderBy) = 0 Then mstrOrderBy = 0
	
	mstrSortOrder = LoadRequestValue("SortOrder")
	If len(mstrSortOrder) = 0 Then mstrSortOrder = "Desc"

	dim paryOrderBy(9)
	paryOrderBy(0) = "orderDate"	'Default
	paryOrderBy(1) = "ssOrderFlagged"
	paryOrderBy(2) = "OrderID"
	paryOrderBy(3) = "custLastName"
	paryOrderBy(4) = "Sum(sfOrderDetails.odrdtQuantity)"
	paryOrderBy(5) = "orderGrandTotal"
	paryOrderBy(6) = "orderDate"
	paryOrderBy(7) = "ssDatePaymentReceived"
	paryOrderBy(8) = "ssDateOrderShipped"
	paryOrderBy(9) = "ssBackOrderDateExpected"

	'WHERE
	'GROUP BY
	'HAVING
	'ORDER
	
	pstrOrderBy = " Order By " & paryOrderBy(mstrOrderBy) & " " & mstrSortOrder 
	'Response.Write "SummaryFilter:<br>WHERE: " & pstrsqlWhere & "<BR>GROUP BY: " & pstrsqlGroupBy & "<BR>HAVING: " & pstrsqlHaving & "<BR>ORDER: "  & pstrOrderBy
	'Response.Flush
	SummaryFilter = Array(pstrsqlWhere, pstrsqlGroupBy, pstrsqlHaving, pstrOrderBy)
	
End Function    'SummaryFilter

'***********************************************************************************************

Sub ShowTemplateSelections

%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblSummaryFunctions">
  <tr>
    <td valign=middle>
		<a href="" id="A1" name=btnShowOrderSummary title="show printable order summaries in a new window" onclick="return btnShowOrderSummary_onclick(this); return false;">Printer Friendly View</a>
		<% If isAllowedToDeleteOrder Then %>&nbsp;|&nbsp;
		<a href="" onclick="DeleteSelected(); return false;">Delete Selected Orders</a>
		<% End If	'isAllowedToDeleteOrder %>
    <!--
		<a href="" id=btnShowOrderSummary name=btnShowOrderSummary title="show printable order summaries in a new window" onclick="return btnShowOrderSummary_onclick(this); return false;">Printer Friendly View</a>
	 -->
    </td>
    <td align="right" valign=middle>
		<select name="ExportTemplates" id="ExportTemplates">
		  <option value="" selected>Select a Template</option>
		<% For i = 0 To UBound(maryExportTemplates) %>
		  <option value="<%= maryExportTemplates(i) %>"><%= Replace(maryExportTemplates(i), ".xsl", "") %></option>
		<% Next 'i %>
		</select>
      <input class="butn" id="btnView" name="btnView" type=button value="View" onclick="viewOrders(''); return false;" title="View Selected Orders">&nbsp;&nbsp;
      <input class="butn" id="btnDownload" name="btnDownload" type=image src="images/save.gif" value="Download" onclick="downloadOrders(''); return false;" title="Download selected orders using the selected template">&nbsp;&nbsp;
      <input class="butn" id="btnPrint" name="btnPrint" type=image src="images/print.gif" value="Print" onclick="printOrders(); return false;" title="Print Selected Orders">&nbsp;&nbsp;
      <% If FileExists(ssAdminPath & "ssOrderAdmin_Export.asp") Then %>
      <br>
      <%
      If Len(cstrPaymentExportFile) > 0 Then
		For i = 0 To UBound(maryExportTemplates)
			If LCase(cstrPaymentExportFile) = LCase(maryExportTemplates(i)) Then
				%><a href="" onclick="return downloadOrders('', '<%= cstrPaymentExportFile %>', 'ssExportedPayment');" title="Create payment export file for the selected orders. Orders are marked as exported."><%= Replace(maryExportTemplates(i), ".xsl", "") %> Export</a>&nbsp;&nbsp;|&nbsp;&nbsp;<%
				Exit For
			End If
		Next 'i
      End If	'

      If Len(cstrShippingExportFile) > 0 Then
		For i = 0 To UBound(maryExportTemplates)
			If LCase(cstrShippingExportFile) = LCase(maryExportTemplates(i)) Then
				%><a href="" onclick="return downloadOrders('','<%= cstrShippingExportFile %>', 'ssExportedShipping');" title="Create shipping export file for the selected orders. Orders are marked as exported."><%= Replace(maryExportTemplates(i), ".xsl", "") %> Export</a>&nbsp;&nbsp;|&nbsp;&nbsp;<%
				Exit For
			End If
		Next 'i
      End If	'
      %>
      <a href="" onclick="return viewOrders('ssOrderAdmin_QuickBooksExport.asp', '', 'ssExported');" title="Create necessary iif files for the selected orders. This marks orders as exported.">Accounting Export</a>&nbsp;&nbsp;
      <% End If	'ssOrderAdmin_QuickBooksExport.asp check %>
      </td>
  </tr>
</table>
<%
End Sub	'ShowTemplateSelections

'***********************************************************************************************

Sub ShowFilter
%>
<script language=javascript>

function quickFilter(theOption)
{
	switch (theOption.value)
	{
		case "0":
			theDataForm.optDate_Filter[0].checked = true;
			ChangeDate(theDataForm.optDate_Filter[0]);
			theDataForm.optText_Filter[0].checked = true;
			theDataForm.optPayment_Filter[0].checked = true;
			theDataForm.optShipment_Filter[0].checked = true;
			theDataForm.optFlag_Filter[2].checked = true;
			theDataForm.Flag_ssExported[2].checked = true;
			theDataForm.Flag_ssExportedPayment[2].checked = true;
			theDataForm.Flag_ssssExportedShipping[2].checked = true;
			theDataForm.Flag_orderVoided[2].checked = true;
			theDataForm.optIncomplete_Filter[0].checked = true;
			
			//theDataForm.optCustomFilter[0].checked = true;
			//theDataForm.optBackOrdered[0].checked = true;
			break;
		case "1":
			theDataForm.optDate_Filter[5].checked = true;
			ChangeDate(theDataForm.optDate_Filter[5]);
			theDataForm.optText_Filter[0].checked = true;
			theDataForm.optPayment_Filter[0].checked = true;
			theDataForm.optShipment_Filter[2].checked = true;
			theDataForm.optFlag_Filter[2].checked = true;
			theDataForm.Flag_ssExported[2].checked = true;
			theDataForm.Flag_ssExportedPayment[2].checked = true;
			theDataForm.Flag_ssssExportedShipping[2].checked = true;
			theDataForm.Flag_orderVoided[2].checked = true;
			theDataForm.optIncomplete_Filter[0].checked = true;
			
			//theDataForm.optCustomFilter[0].checked = true;
			//theDataForm.optBackOrdered[0].checked = true;
			break;
		case "2":
			theDataForm.optDate_Filter[5].checked = true;
			ChangeDate(theDataForm.optDate_Filter[5]);
			theDataForm.optText_Filter[0].checked = true;
			theDataForm.optPayment_Filter[1].checked = true;
			theDataForm.optShipment_Filter[0].checked = true;
			theDataForm.optFlag_Filter[2].checked = true;
			theDataForm.Flag_ssExported[2].checked = true;
			theDataForm.Flag_ssExportedPayment[2].checked = true;
			theDataForm.Flag_ssssExportedShipping[2].checked = true;
			theDataForm.Flag_orderVoided[2].checked = true;
			theDataForm.optIncomplete_Filter[0].checked = true;
			
			//theDataForm.optCustomFilter[0].checked = true;
			//theDataForm.optBackOrdered[0].checked = true;
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "3":
			theDataForm.optDate_Filter[5].checked = true;
			ChangeDate(theDataForm.optDate_Filter[5]);
			theDataForm.optText_Filter[0].checked = true;
			theDataForm.optPayment_Filter[2].checked = true;
			theDataForm.optShipment_Filter[2].checked = true;
			theDataForm.optFlag_Filter[2].checked = true;
			theDataForm.Flag_ssExported[2].checked = true;
			theDataForm.Flag_ssExportedPayment[2].checked = true;
			theDataForm.Flag_ssssExportedShipping[2].checked = true;
			theDataForm.Flag_orderVoided[2].checked = true;
			theDataForm.optIncomplete_Filter[1].checked = true;
			
			//theDataForm.optCustomFilter[0].checked = true;
			//theDataForm.optBackOrdered[0].checked = true;
			//theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
	}

}
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <tr>
    <td>
     Order Views:&nbsp;<input type="radio" value="0" <% if mbytDisplay=enDisplayType_OrderDetail then Response.Write "Checked" %> name="optDisplay" ID="optDisplay0"><span onclick="theDataForm.optDisplay[0].checked=true;">Detail&nbsp;</span>
      &nbsp;<input type="radio" value="1" <% if mbytDisplay=enDisplayType_OutstandingPayments then Response.Write "Checked" %> name="optDisplay" onclick="ChangeDisplay(true);" ID="optDisplay1"><span onclick="theDataForm.optDisplay[1].checked=true;ChangeDisplay(true);">Quick&nbsp;Payment</span>
      &nbsp;<input type="radio" value="2" <% if mbytDisplay=enDisplayType_OutStandingShipments then Response.Write "Checked" %> name="optDisplay" onclick="ChangeDisplay(false);" ID="optDisplay2"><span onclick="theDataForm.optDisplay[2].checked=true;ChangeDisplay(false);">Quick&nbsp;Shipment</span>
      &nbsp;<input type="radio" value="3" <% if mbytDisplay=enDisplayType_ImportTracking then Response.Write "Checked" %> name="optDisplay" onclick="ChangeDisplay(false);" ID="optDisplay3"><span onclick="theDataForm.optDisplay[3].checked=true;ChangeDisplay(false);">Import&nbsp;Tracking</span><br>

     Saved Filters:&nbsp;<input type="radio" value="0" <% if mbytQuick_Filter="0" then Response.Write "Checked" %> name="optQuick_Filter" ID="optQuick_Filter0" onclick="quickFilter(this);"><label for="optQuick_Filter0">Today's&nbsp;Orders</label>
      &nbsp;<input type="radio" value="1" <% if mbytQuick_Filter="1" then Response.Write "Checked" %> name="optQuick_Filter" ID="optQuick_Filter1" onclick="quickFilter(this);"><label for="optQuick_Filter1">Orders&nbsp;Awaiting&nbsp;Payment</label>
      &nbsp;<input type="radio" value="2" <% if mbytQuick_Filter="2" then Response.Write "Checked" %> name="optQuick_Filter" ID="optQuick_Filter2" onclick="quickFilter(this);"><label for="optQuick_Filter2">Orders&nbsp;Awaiting&nbsp;Shipment</label>
      &nbsp;<input type="radio" value="3" <% if mbytQuick_Filter="3" then Response.Write "Checked" %> name="optQuick_Filter" ID="optQuick_Filter3" onclick="quickFilter(this);"><label for="optQuick_Filter3">Incomplete&nbsp;Orders</label>
    </td>
    <td>
	  <p align="right"><INPUT class="butn" id="btnFilter" name="btnFilter" type=button value="Apply Filter" onclick="btnFilter_onclick(this);">&nbsp;&nbsp;</p>
    </td>
  </tr>
  <tr>
    <td valign="top">
	  <fieldset>
	    <legend>Orders containing text in field:</legend>
        <input type="text" id="Text_Filter" name="Text_Filter" size="20" value="<%= Server.HTMLEncode(mstrText_Filter) %>"><br />
        <input type="radio" value="0" <% if (mbytText_Filter="0" or mbytText_Filter="") then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter0"><label for="optText_Filter0">Do Not Include</label><br />
		<% For i = 1 To UBound(maryTextFilterOptions) %>
			<input type="radio" value="<%= i %>" <% if mbytText_Filter=CStr(i) then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter<%= i %>"><label for="optText_Filter<%= i %>"><%= maryTextFilterOptions(i) %></label><br />
		<% Next 'i %>
	  </fieldset>
	  <p>
        </p>
        
      <fieldset>
        <legend>Show orders products by:</legend>
        <p>Specific Products<br>
		<select id="ProductID" name="ProductID" size="5" ondblclick="openMovementWindow('ProductID','product');" multiple></select>
		<a href="" onclick="openMovementWindow('ProductID','product'); return false;"><img src="images/properites.gif" border="0"></a>
		</p>
      
		<p>From Manufacturer:<br>
		<select name="Mfg_Filter" id="Mfg_Filter">
		<% If Len(mstrMfg_Filter) > 0 Then %>
		<option value="" selected>Ignore</option>
		<% Else %>
		<option value="">Ignore</option>
		<% End If %>
		<%= createCombo("SELECT Distinct odrdtManufacturer FROM sfOrderDetails ORDER BY odrdtManufacturer", "", "odrdtManufacturer", mstrMfg_Filter) %>
		</select></p>
		
		<p>From Vendor:<br>
		<select name="Vend_Filter" id="Vend_Filter">
		<% If Len(mstrVend_Filter) > 0 Then %>
		<option value="" selected>Ignore</option>
		<% Else %>
		<option value="">Ignore</option>
		<% End If %>
		<%= createCombo("SELECT Distinct odrdtVendor FROM sfOrderDetails ORDER BY odrdtVendor", "", "odrdtVendor", mstrVend_Filter) %>
		</select></p>
      </fieldset>

	</td>
    <td valign="top">
      <fieldset>
        <legend>Show Only Orders Placed</legend>
        <input type="radio" value="1" <% if mbytDate_Filter="1" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter1"><label for="optDate_Filter1">Day</label>&nbsp;
        <input type="radio" value="2" <% if mbytDate_Filter="2" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter2"><label for="optDate_Filter2">Week</label>&nbsp;
        <input type="radio" value="3" <% if mbytDate_Filter="3" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter3"><label for="optDate_Filter3">Month</label>&nbsp;
        <input type="radio" value="4" <% if mbytDate_Filter="4" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter4"><label for="optDate_Filter4">Year</label>&nbsp;
        <input type="radio" value="5" <% if mbytDate_Filter="5" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter5"><label for="optDate_Filter5">Custom</label>&nbsp;
        <input type="radio" value="0" <% if mbytDate_Filter="0" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter0"><label for="optDate_Filter0">All</label>&nbsp;<br>
		<label for="StartDate">Start Date:&nbsp;</label><input id=StartDate name=StartDate Value="<%= mstrStartDate %>">
		<a HREF="javascript:doNothing()" title="Select start date"
		onClick="setDateField(document.frmData.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER=0></a><br>

		<label for="EndDate">&nbsp;&nbsp;End Date:&nbsp;</label><input id=EndDate name=EndDate Value="<%= mstrEndDate %>">
		<a HREF="javascript:doNothing()" title="Select end date"
		onClick="setDateField(document.frmData.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER=0></a><br />
		Using&nbsp;
		<% For i = 0 To UBound(maryDateType_Filter) %>
        <input type="radio" value="<%= i %>" <% if mbytDateType_Filter=CStr(i) then Response.Write "Checked" %> name="optDateType_Filter" id="optDateType_Filter<%= i %>"><label for="optDateType_Filter<%= i %>"><%= maryDateType_Filter(i)(0) %></label>&nbsp;
		<% Next 'i %>
      </fieldset>

	  <fieldset>
	    <legend>Show Orders that are:</legend>
		<table class="tbl" cellpadding="0" cellspacing="0" border="0" ID="Table2">
		  <tr>
		    <td><input type="radio" value="2" <% if mbytPayment_Filter="2" then Response.Write "Checked" %> name="optPayment_Filter" ID="optPayment_Filter2"><label for="optPayment_Filter2"><font size="-1">Awaiting&nbsp;Payment</font></label></td>
		    <td>&nbsp;<input type="radio" value="1" <% if mbytPayment_Filter="1" then Response.Write "Checked" %> name="optPayment_Filter" ID="optPayment_Filter1"><label for="optPayment_Filter1"><font size="-1">Paid</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytPayment_Filter="0" or mbytPayment_Filter="") then Response.Write "Checked" %> name="optPayment_Filter" ID="optPayment_Filter0"><label for="optPayment_Filter0"><font size="-1">All</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytShipment_Filter="1" then Response.Write "Checked" %> name="optShipment_Filter" ID="optShipment_Filter1"><label for="optShipment_Filter1"><font size="-1">Awaiting&nbsp;Shipment</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytShipment_Filter="2" then Response.Write "Checked" %> name="optShipment_Filter" ID="optShipment_Filter2"><label for="optShipment_Filter2"><font size="-1">Shipped</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytShipment_Filter="0" or mbytShipment_Filter="") then Response.Write "Checked" %> name="optShipment_Filter" ID="optShipment_Filter0"><label for="optShipment_Filter0"><font size="-1">All</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytoptFlag_Filter="1" then Response.Write "Checked" %> name="optFlag_Filter" ID="optFlag_Filter1"><label for="optFlag_Filter1"><font size="-1">Flagged</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytoptFlag_Filter="2" then Response.Write "Checked" %> name="optFlag_Filter" ID="optFlag_Filter2"><label for="optFlag_Filter2"><font size="-1">Unflagged</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytoptFlag_Filter="0" or mbytoptFlag_Filter="") then Response.Write "Checked" %> name="optFlag_Filter" ID="optFlag_Filter0"><label for="optFlag_Filter0"><font size="-1">Both</font></label></td>
		  </tr>
		  <% If cblnUseBackOrder Then %>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytoptBackOrdered="1" then Response.Write "Checked" %> name="optBackOrdered" ID="optBackOrdered1"><label for="optBackOrdered1"><font size="-1">Back&nbsp;Ordered</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytoptBackOrdered="2" then Response.Write "Checked" %> name="optBackOrdered" ID="optBackOrdered2"><label for="optBackOrdered2"><font size="-1">Not&nbsp;BackOrdered</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytoptBackOrdered="0" or mbytoptBackOrdered="") then Response.Write "Checked" %> name="optBackOrdered" ID="optBackOrdered0"><label for="optBackOrdered0"><font size="-1">Both</font></label></td>
		  </tr>
		  <% End If %>
		  <tr>
		    <td><input type="radio" value="0" <% if (mbytIncomplete_Filter="0" or mbytIncomplete_Filter="") then Response.Write "Checked" %> name="optIncomplete_Filter" ID="optIncomplete_Filter0"><label for="optIncomplete_Filter0"><font size="-1">Complete</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytIncomplete_Filter="2" then Response.Write "Checked" %> name="optIncomplete_Filter" ID="optIncomplete_Filter2"><label for="optIncomplete_Filter2"><font size="-1">Incomplete</font></label></td>
		    <td>&nbsp;<input type="radio" value="1" <% if mbytIncomplete_Filter="1" then Response.Write "Checked" %> name="optIncomplete_Filter" ID="optIncomplete_Filter1"><label for="optIncomplete_Filter1"><font size="-1">All</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytFlag_orderVoided="1" then Response.Write "Checked" %> name="Flag_orderVoided" id="Flag_orderVoided1"><label for="Flag_orderVoided1"><font size="-1">Voided</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytFlag_orderVoided="2" then Response.Write "Checked" %> name="Flag_orderVoided" id="Flag_orderVoided2"><label for="Flag_orderVoided2"><font size="-1">Not&nbsp;Voided</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytFlag_orderVoided="0" or mbytFlag_orderVoided="") then Response.Write "Checked" %> name="Flag_orderVoided" id="Flag_orderVoided0"><label for="Flag_orderVoided0"><font size="-1">Both</font></label></td>
		  </tr>
		</table>
	  </fieldset>
	  <fieldset>
	    <legend>Show Orders that are:</legend>
		<table class="tbl" cellpadding="0" cellspacing="0" border="0" ID="Table1">
		  <tr>
		    <td><input type="radio" value="1" <%= isChecked(mbytFlag_ssExported="1") %> name="Flag_ssExported" ID="Flag_ssExported1"><label for="Flag_ssExported1"><font size="-1">Exported to Accounting</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <%= isChecked(mbytFlag_ssExported="2") %> name="Flag_ssExported" ID="Flag_ssExported2"><label for="Flag_ssExported2"><font size="-1">Not&nbsp;Exported</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <%= isChecked(mbytFlag_ssExported="0" or mbytFlag_ssExported="") %> name="Flag_ssExported" ID="Flag_ssExported0"><label for="Flag_ssExported0"><font size="-1">Both</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <%= isChecked(mbytFlag_ssExportedPayment="1") %> name="Flag_ssExportedPayment" ID="Flag_ssExportedPayment1"><label for="Flag_ssExportedPayment1"><font size="-1">Exported to Payment</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <%= isChecked(mbytFlag_ssExportedPayment="2") %> name="Flag_ssExportedPayment" ID="Flag_ssExportedPayment2"><label for="Flag_ssExportedPayment2"><font size="-1">Not&nbsp;Exported</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <%= isChecked(mbytFlag_ssExportedPayment="0" or mbytFlag_ssExportedPayment="") %> name="Flag_ssExportedPayment" ID="Flag_ssExportedPayment"><label for="Flag_ssExportedPayment0"><font size="-1">Both</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <%= isChecked(mbytFlag_ssssExportedShipping="1") %> name="Flag_ssssExportedShipping" ID="Flag_ssssExportedShipping1"><label for="Flag_ssssExportedShipping1"><font size="-1">Exported to Shipping</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <%= isChecked(mbytFlag_ssssExportedShipping="2") %> name="Flag_ssssExportedShipping" ID="Flag_ssssExportedShipping2"><label for="Flag_ssssExportedShipping2"><font size="-1">Not&nbsp;Exported</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <%= isChecked(mbytFlag_ssssExportedShipping="0" or mbytFlag_ssssExportedShipping="") %> name="Flag_ssssExportedShipping" ID="Flag_ssssExportedShipping0"><label for="Flag_ssssExportedShipping0"><font size="-1">Both</font></label></td>
		  </tr>
		</table>
	  </fieldset>
	  
	  <fieldset>
	    <legend>Show Only Orders with</legend>
		<p>External order status of:<br>
		<select name="Flag_OrderStatus" id="Flag_OrderStatus">
		<% If CLng(mbytOrderStatus_Filter) = -1 Then %>
		<option value="-1" selected>Ignore</option>
		<% Else %>
		<option value="-1">Ignore</option>
		<% End If %>
		<% For i = 0 To UBound(maryOrderStatuses) %>
		<% If i = CLng(mbytOrderStatus_Filter) Then %>
		<option value="<%= i %>" selected><%= maryOrderStatuses(i) %></option>
		<% Else %>
		<option value="<%= i %>"><%= maryOrderStatuses(i) %></option>
		<% End If %>
		<% Next 'i %>
		</select></p>
		
		<p>Internal order status of:<br>
		<select name="Flag_InternalOrderStatus" id="Flag_InternalOrderStatus">
		<% If CLng(mbytInternalOrderStatus_Filter) = -1 Then %>
		<option value="-1" selected>Ignore</option>
		<% Else %>
		<option value="-1">Ignore</option>
		<% End If %>
		<% For i = 0 To UBound(maryInternalOrderStatuses) %>
		<% If i = CLng(mbytInternalOrderStatus_Filter) Then %>
		<option value="<%= i %>" selected><%= maryInternalOrderStatuses(i)(0) %></option>
		<% Else %>
		<option value="<%= i %>"><%= maryInternalOrderStatuses(i)(0) %></option>
		<% End If %>
		<% Next 'i %>
		</select></p>
		
		<p>Payment type of:<br>
		<select name="optPaymentType" id="optPaymentType">
		<% If Len(mstrPaymentType_Filter) > 0 Then %>
		<option value="" selected>Ignore</option>
		<% Else %>
		<option value="">Ignore</option>
		<% End If %>
		<% If mstrPaymentType_Filter = "zilch" Then %>
		<option value="zilch" selected>Missing</option>
		<% mstrPaymentType_Filter = ""	'added so this doesn't appear in combo %>
		<% Else %>
		<option value="zilch">Missing</option>
		<% End If %>
		<%= createCombo("SELECT Distinct orderPaymentMethod FROM sfOrders ORDER BY orderPaymentMethod", "", "orderPaymentMethod", mstrPaymentType_Filter) %>
		</select></p>
	  </fieldset>
	  <% If cblnUseCustomFilter Then %>
	  <fieldset>
		<legend>Show Orders From Store</legend>
		<input type="radio" value="0" <% if mbytCustomFilter="0" then Response.Write "Checked" %> name="optCustomFilter" ID="optCustomFilter0"><label for="optCustomFilter0">All</label><br />
		<input type="radio" value="1" <% if mbytCustomFilter="1" then Response.Write "Checked" %> name="optCustomFilter" ID="optCustomFilter1"><label for="optCustomFilter1">SMD</label><br />
		<input type="radio" value="2" <% if mbytCustomFilter="2" then Response.Write "Checked" %> name="optCustomFilter" ID="optCustomFilter2"><label for="optCustomFilter2">MEU</label><br />
	  </fieldset>
	  <% End If %>
		
	</td>
  </tr>
</table>
<% End Sub	'ShowFilter %>
