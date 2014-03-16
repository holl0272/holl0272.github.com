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

'here's what to search for when FrontPage corrupts the file
'Search for		: =>
'Replace with	: ="
'Search for		: mailto><%=
'Replace with	: mailto:<%=
'Search for		: ID><%=
'Replace with	: ID=<%=
'

'***********************************************************************************************

Dim cblnUseCustomOrderFields:	cblnUseCustomOrderFields = False	'Placeholder
Dim maryCustomDisplayColumns
Dim enCustomColumns_FieldName:			enCustomColumns_FieldName = 0				' 
Dim enCustomColumns_ColumnName:			enCustomColumns_ColumnName = 1			' 
Dim enCustomColumns_ColumnWidth:		enCustomColumns_ColumnWidth = 2		' 
Dim enCustomColumns_ColumnAlignment:	enCustomColumns_ColumnAlignment = 3	' 

'***********************************************************************************************

Sub InitializeCustomColumns

	Exit Sub

	ReDim maryCustomDisplayColumns(1)

	maryCustomDisplayColumns(0) = Array("prodManufacturerId", "Mfg ID", "", "center")
	maryCustomDisplayColumns(1) = Array("prodVendorId", "Vendor ID", "", "center")
	
	'maryCustomDisplayColumns(0) = Array("prodMFGNumber", "Mfg ID", "", "center")
	'maryCustomDisplayColumns(1) = Array("prodVenderNumber", "Vendor ID", "", "center")

End Sub

'***********************************************************************************************

Function orderDetailColumnCount(byVal NumCellsBefore, byVal NumCellsAfter)

Dim plngOrderDetailColumnCount

	plngOrderDetailColumnCount = 6	'set the base number
	
	If isArray(maryCustomDisplayColumns) Then plngOrderDetailColumnCount = plngOrderDetailColumnCount + UBound(maryCustomDisplayColumns) + 1
	If cblnShowBackOrderColumn Then plngOrderDetailColumnCount = plngOrderDetailColumnCount  + 1
	
	orderDetailColumnCount = plngOrderDetailColumnCount - NumCellsBefore- NumCellsAfter

End Function	'orderDetailColumnCount

'***********************************************************************************************

Sub writeEmptyCells(byVal colSpan)
	Response.Write "<td colspan=" & colSpan & ">&nbsp;</td>"
End Sub

'***********************************************************************************************

Sub ShowOrderDetail(objRS)

Dim plngOrderID
Dim pcurRealSubTotal
Dim pstrCustName
Dim pstrCustAddr
Dim pstrProdID
Dim pblnOddRow
Dim pstrBackground
Dim pstrTempID
Dim plngodrdtID
Dim pstrInventoryMessage
Dim pstrAttributeName
Dim pstrAttributeDetailName
Dim pstrProductName
Dim i
Dim plngNumberDetailColumns

	If Not isObject(objRS) Then Exit Sub
	
	plngOrderID = objRS.Fields("orderID").Value
	pblnOddRow = False

	pstrCustName = Replace(objRS.Fields("custFirstName").Value & " " & objRS.Fields("custMiddleInitial").Value & " " & objRS.Fields("custLastName").Value,"  "," ")
	pstrCustAddr = objRS.Fields("custCity").Value & ", " & objRS.Fields("custState").Value & " " & objRS.Fields("custZip").Value

	Call InitializeCustomColumns

	If cblnSF5AE And Not cblnShowBackOrderColumn Then
		Do While Not objRS.EOF
			If objRS.Fields("odrdtBackOrderQTY").Value > 0 Then cblnShowBackOrderColumn = True
			objRS.MoveNext
		Loop
		objRS.MoveFirst
	End If

%>
<input type="hidden" name="quantityUpdated" id="quantityUpdated" value="0">
<table class="tbl" width="100%" cellpadding="0" cellspacing="0" border="0" rules="none" id="tblOrderDetail">
  <tr>
    <td width="100%" colspan="2">
      <table class="tbl" style="border-collapse: collapse" border="0" cellspacing="0" cellpadding="1" width="100%" ID="tblOrderDetailSummary">
		<colgroup>
		  <col valign="top" /><!-- Checkbox -->
		  <col valign="top" />
		  <col valign="top" />
		  <col valign="top" />
		  <col valign="top" />
		  <col valign="top" />
		  <col valign="top" />
		</colgroup>
		<% 
		If isArray(maryCustomDisplayColumns) Then
			For i = 0 To UBound(maryCustomDisplayColumns)
		%>
		<colgroup align="<%= maryCustomDisplayColumns(i)(enCustomColumns_ColumnAlignment) %>" width="<%= maryCustomDisplayColumns(i)(enCustomColumns_ColumnAlignment) %>"><!-- Custom Field -->
		<%
			Next 'i
		End If
		%>
		<colgroup align=left><!-- Product Name -->
		<colgroup align=center><!-- Quantity -->
		<% If cblnShowBackOrderColumn Then %><colgroup align=center><!-- Back order --><% End If %>
		<colgroup align=center><!-- Unit Price -->
		<colgroup align=center><!-- Extended Price -->
        <tr>
          <th align=center>&nbsp;</th>
          <th align=center>Product ID</th>
		<% 
		If isArray(maryCustomDisplayColumns) Then
			For i = 0 To UBound(maryCustomDisplayColumns)
		%>
          <th align=center><%= maryCustomDisplayColumns(i)(enCustomColumns_ColumnName) %></th>
		<%
			Next 'i
		End If
		%>
          <th align=center>Product Name</th>
          <th align=center>Quantity</th>
          <% If cblnShowBackOrderColumn Then %><th align=center>Qty Backordered</th><% End If %>
          <th align=center>Unit Price</th>
          <th align=center>Extended Price</th>
          <th align=center>Points</th>
        </tr>
<%
Dim p_strProdIDLink
Dim pblnWriteGiftWrapBackorder
Dim pblnStepBack

Dim cblnDisplayGiftWrap
Dim pstrGWQty
Dim pstrGWPrice
Dim pUnitPrice
Dim pLineItemPrice
cblnDisplayGiftWrap = True

pcurRealSubTotal = 0
pblnWriteGiftWrapBackorder = False
pblnStepBack = False

Do While Not objRS.EOF

	If plngodrdtID <> Trim(objRS.Fields("odrdtID").Value) Then
		plngodrdtID = Trim(objRS.Fields("odrdtID").Value)
		pstrProdID = Trim(objRS.Fields("odrdtProductID").Value)
		pstrProductName = Replace(Trim(objRS.Fields("odrdtProductName").Value & ""), "," , "")
		
		If cblnSF5AE Then
			If Len(objRS.Fields("odrdtGiftWrapQTY").Value) > 0 Then pstrGWQty = CLng(objRS.Fields("odrdtGiftWrapQTY").Value)
			If pstrGWQty > 0 Then
				pstrGWPrice = objRS.Fields("odrdtGiftWrapPrice").Value/pstrGWQty
			Else
				pstrGWQty = 0
				pstrGWPrice = 0
			End If
			
			If cblnShowInventoryOnHand Then
				Dim plngInventoryQty
				Dim plngInventoryLow
				Dim pstrodrdtAttDetailID
				
				pstrodrdtAttDetailID = objRS.Fields("odrdtAttDetailID")
				If isInventoryTracked(pstrProdID) Then
					plngInventoryQty = inventoryQty(pstrProdID, pstrodrdtAttDetailID)
					plngInventoryLow = inventoryLowQty(pstrProdID, pstrodrdtAttDetailID)
					If plngInventoryQty < plngInventoryLow Then
						pstrInventoryMessage = "&nbsp;<span style='color:red;font-weight: bold;' title='This product is below the notification quantity of " & plngInventoryLow & "'>(" & plngInventoryQty & ")</span>"
					Else
						pstrInventoryMessage = "&nbsp;<span style='color:black;'>(" & plngInventoryQty & ")</span>"
					End If
				Else
					pstrInventoryMessage = ""
					pstrInventoryMessage = "&nbsp;<span style='color:black;'>(-)</span>"
				End If
				pstrodrdtAttDetailID = Replace(pstrodrdtAttDetailID & "", ",", "|")	'Prepare for insertion into hidden field
			End If

		End If
		
		pblnOddRow = Not pblnOddRow
		If pblnOddRow Then
			pstrBackground = " style=""background: lightgrey;"""	'set this color for the odd rows
		Else
			pstrBackground = " style=""background: white;"""				'set this color for the even rows
		End If

		If cblnAddon_ProductMgr Then
			p_strProdIDLink = "<a href=" & Chr(34) & "sfProductAdmin.asp?Action=Deactivate&prodID=" & Server.URLEncode(pstrProdID) & "&radTextSearch=3&TextSearch=" & Server.URLEncode(pstrProductName) & Chr(34) & " title='Click to deactivate this product using Product Manager and apply a filter looking for other products with the same name in the long description' target='ProductManager'>-</a>"
			'p_strProdIDLink = p_strProdIDLink & "&nbsp;&nbsp;" _
			'				& "<a href=" & Chr(34) & "sfProductAdmin.asp?Action=ViewProduct&radTextSearch=3&ViewID=" & Server.URLEncode(pstrProdID) & "&TextSearch=" & Server.URLEncode(objRS.Fields("odrdtProductName").Value) & Chr(34) & " title='Click to view using Product Manager' target='ProductManager'>" & pstrProdID & "</a>"

			p_strProdIDLink = p_strProdIDLink & "&nbsp;&nbsp;" _
							& "<a href=" & Chr(34) & "sfProductAdmin.asp?Action=ViewProduct&radTextSearch=1&ViewID=" & Server.URLEncode(pstrProdID) & "&TextSearch=" & Server.URLEncode(pstrProdID) & Chr(34) & " title='Click to view using Product Manager' target='ProductManager'><img width='16' height='16' src='images/PREVIEW.BMP' alt='preview' title='Click to view using Product Manager'></a>"
		Else
			p_strProdIDLink = pstrProdID
		End If
		
		pUnitPrice = Trim(objRS.Fields("odrdtSubTotal").Value & "")
		If Len(pUnitPrice) > 0 And isNumeric(pUnitPrice) Then
			pUnitPrice = CDbl(pUnitPrice)
		Else
			Response.Write "<font color=red>Extended price of (" & pUnitPrice & ") is not a number. It has been reset to zero.</font>"
			pUnitPrice = 0
		End If
		If objRS.Fields("odrdtQuantity").Value <> 0 Then
			pLineItemPrice = pUnitPrice
			pUnitPrice = pUnitPrice / objRS.Fields("odrdtQuantity").Value
		Else
			pLineItemPrice = 0
		End If
		
%>
        <tr <%= pstrBackground %>" id="sampleProductRow">
          <td>
            <input type="hidden" name="odrdtID" id="odrdtID" value="<%= plngodrdtID %>">
            <input type="hidden" name="origQty" id="origQty" value="<%= objRS.Fields("odrdtQuantity").Value %>">
			<input type="hidden" name="odrdtAttDetailID" id="odrdtAttDetailID" value="<%= pstrodrdtAttDetailID %>">      
			<input type="checkbox" name="deleteodrdtID" id="deleteodrdtID" value="<%= plngodrdtID %>" onclick="deleteOrderDetail(this);" title="Remove this item from the order">&nbsp;         
            <%= p_strProdIDLink %>
          </td>
          <td>
            <input name="odrdtProductID" id="odrdtProductID.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:left;" onchange="makeBaseOrderDirty();" value="<%= pstrProdID %>" size="<%= Len(pstrProdID) + 2 %>">
<%
'added for download management
Dim pblnDownloadAvailable
Dim pstrDownloadMessage

pblnDownloadAvailable = HasDownloadAvailable_orderDetail(objRS.Fields("orderCustId").Value, plngodrdtID)
If Not pblnDownloadAvailable Then pblnDownloadAvailable = CBool(Download_RequestStatus <> enDownloadRequest_NoDownloadAvailable) And Not CBool(Download_RequestStatus = enDownloadRequest_InvalidRequest)

pstrDownloadMessage = DownloadRequest_RequestStatusText(Download_RequestStatus)
If pblnDownloadAvailable Then
%>
<fieldset><legend>Downloadable Item</legend>
          <table border=0 cellpadding=1 cellspacing=0>
          <tr>
          <td><label for="">Status:</label></td>
          <td><%= pstrDownloadMessage %></td>
          </tr>
          <tr>
          <td><label for="odrdtDownloadExpiresOn.<%= plngodrdtID %>" title="Last downloaded <%= Download_LastDownload %>">Expires:</label></td>
          <td><input type="text" name="odrdtDownloadExpiresOn.<%= plngodrdtID %>" id="odrdtDownloadExpiresOn.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:left;" onchange="makeBaseOrderDirty();" value="<%= Download_ExpiresOn %>" size="10"></td>
          </tr>
          <tr>
          <td><label for="">Downloads:</label></td>
          <td><%= Download_CurrentDownloadCount %> of <input type="text" name="odrdtMaxDownloads.<%= plngodrdtID %>" id="odrdtMaxDownloads.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:left;" onchange="makeBaseOrderDirty();" ondblclick="var count; if (this.value.length > 0){count = new Number(this.value);}else{count=0;} count=count+1; this.value=count.toString();" value="<%= Download_MaxDownloads %>" size="3"></td>
          </tr>
          <tr>
          <td><label for="">Authorized:</label></td>
          <td><input type="checkbox" name="odrdtDownloadAuthorized.<%= plngodrdtID %>" id="odrdtDownloadAuthorized.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" onclick="makeBaseOrderDirty();" value="1" <%= isChecked(objRS.Fields("odrdtDownloadAuthorized").Value) %>></td>
          </tr>
          </table>
</fieldset>
<%
ElseIf Download_RequestStatus <> enDownloadRequest_NoDownloadAvailable Then
	Response.Write pstrDownloadMessage & " (" & Download_RequestStatus & ")"
End If
%>
          </td>
		<% 
		If isArray(maryCustomDisplayColumns) Then
			For i = 0 To UBound(maryCustomDisplayColumns)
		%>
          <td><%= objRS.Fields(maryCustomDisplayColumns(i)(enCustomColumns_FieldName)).Value %></td>
		<%
			Next 'i
		End If
		%>
          <td>
          <% 'added to display textarea or text box
          If clngProductNameLength = 0 Or clngProductNameLength > Len(pstrProductName) Then
          %>
          <input name="odrdtProductName" id="odrdtProductName.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:left;" onchange="makeBaseOrderDirty();" value="<%= pstrProductName %>" size="<%= Len(pstrProductName) + 2 %>">
          <% Else %>
          <textarea name="odrdtProductName" id="odrdtProductName.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:left;" onchange="makeBaseOrderDirty();" cols="<%= clngProductNameLength %>" rows="<%= Len(pstrProductName)/clngProductNameLength + 1 %>"><%= pstrProductName %></textarea>
          <% End If %>
          </td>
          <td><input name="odrdtQuantity" id="odrdtQuantity.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:center;" onchange="return recalcProductTotal(this); this.form.quantityUpdated.value=1;" value="<%= objRS.Fields("odrdtQuantity").Value %>" size="5"><%= pstrInventoryMessage %></td>
          <% If cblnShowBackOrderColumn Then %><td><input name="odrdtBackOrderQTY" id="odrdtBackOrderQTY.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:center;" onchange="makeBaseOrderDirty();" value="<%= objRS.Fields("odrdtBackOrderQTY").Value %>" size="5"></td><% End If %>
          <td><input name="unitPrice" id="unitPrice.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:right;" onchange="return recalcProductTotal(this);" value="<%= pUnitPrice %>" size="5"></td>
          <td><input name="odrdtSubTotal" id="odrdtSubTotal.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:right;" onchange="return recalcProductTotal(this);" value="<%= Trim(objRS.Fields("odrdtSubTotal").Value & "") %>" size="5"></td>
          <td><input name="buyersClubPointsIssued" id="buyersClubPointsIssued.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:right;" onchange="return recalcProductTotal(this);" value="<%= Trim(objRS.Fields("buyersClubPointsIssued").Value & "") %>" size="5"></td>
        </tr>
<%
		pcurRealSubTotal = pcurRealSubTotal + pLineItemPrice
	End If
	
	pstrAttributeName = Trim(objRS.Fields("odrattrName").Value & " ")
	If Len(pstrAttributeName) > 0 Then
		pstrAttributeDetailName = Trim(objRS.Fields("odrattrAttribute").Value)

		If InStr(1, pstrAttributeName, Left(pstrAttributeDetailName, Len(pstrAttributeName))) > 0 Then 
			pstrAttributeDetailName = Replace(pstrAttributeDetailName, pstrAttributeName, "", 1, 1)
		End If
		
		Dim paryAttributeDetails
		ReDim paryAttributeDetails(1)
		Call getProductAttributeDetails(pstrProdID, pstrAttributeName, pstrAttributeDetailName, Array("attrdtExtra", "attrdtExtra1"), paryAttributeDetails)
		
		'Remove the trailing semi-colon, if present
		If Right(pstrAttributeName, 1) = ":" Then pstrAttributeName = Left(pstrAttributeName, Len(pstrAttributeName) - 1)
%>
        <tr <%= pstrBackground %>>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
		<% 
		If isArray(maryCustomDisplayColumns) Then
			For i = 0 To UBound(maryCustomDisplayColumns)
		%>
          <td><%= paryAttributeDetails(i) %></td>
		<%
			Next 'i
		End If
		%>
          <td><font size="-1">&nbsp;&nbsp;<%= pstrAttributeName %>: <%= pstrAttributeDetailName %></font></td>
          <td>&nbsp;</td>
			<% If cblnShowBackOrderColumn Then %><td>&nbsp;</td><% End If %>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
<%	
	End If

	objRS.MoveNext
	
	If cblnSF5AE Then
		If objRS.EOF Then
			pblnWriteGiftWrapBackorder = True
		Else
			pblnWriteGiftWrapBackorder = CBool(plngodrdtID <> Trim(objRS.Fields("odrdtID").Value))
		End If
	End If

	If pblnWriteGiftWrapBackorder Then
		'Check for Gift Wrap
		If objRS.EOF Then
			pblnStepBack = True
			objRS.MovePrevious
		End If
		
		%>
		<input type="hidden" name="odrdtAEID" id="odrdtAEID" value="<%= plngodrdtID %>">
		<%
		If pstrGWQty > 0 Or cblnShowGiftWrapRow Then 
			pcurRealSubTotal = pcurRealSubTotal + pstrGWPrice * pstrGWQty
		%>
			<tr <%= pstrBackground %>>
			  <td><input type="checkbox" name="deleteodrdtAEID" id="deleteodrdtAEID" value="<%= plngodrdtID %>" onclick="deleteOrderDetailAE(this);" title="Remove gift wrap/back order from this item"></td>
			  <td></td>
			  <td>&nbsp;<i>Gift Wrap</i></td>
			  <td><input name="odrdtGiftWrapQTY" id="odrdtGiftWrapQTY.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:center;" onchange="return recalcGiftWrapTotal(this);" value="<%= pstrGWQty %>" size="5"></td>
			<% If cblnShowBackOrderColumn Then %><td>&nbsp;</td><% End If %>
			  <td><input name="GiftWrapUnitPrice" id="GiftWrapUnitPrice.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:right;" onchange="return recalcGiftWrapTotal(this);" value="<%= pstrGWPrice %>" size="5"></td>
			  <td><input name="odrdtGiftWrapPrice" id="odrdtGiftWrapPrice.<%= plngodrdtID %>" tag="<%= plngodrdtID %>" style="text-align:right;" onchange="return recalcGiftWrapTotal(this);" value="<%= pstrGWPrice * pstrGWQty %>" size="5"></td>
			  <td>&nbsp;</td>
			</tr>
<%			
		End If

		If pblnStepBack Then objRS.MoveNext
	End If
	pblnWriteGiftWrapBackorder = False
	
Loop
objRS.MoveFirst
%>
        <tr>
          <td width="100%" colspan="7" style="border-bottom: black 1pt solid;">&nbsp;</td>
        </tr>
        <tr>
          <td colspan=2 align=left valign=top>
            <input type="button" class="butn" name="btnaddProduct" id="btnaddProduct" value="Add Item" onclick="addProduct(this);"><br>
            <input type="button" class="butn" name="btndeleteProduct" id="btndeleteProduct" value="Remove Item" onclick="deleteProduct(this);"><br>
            <!-- <input type="button" class="butn" name="btndeleteGiftWrap" id="btndeleteGiftWrap" value="Delete Gift Wrap" onclick="deleteGiftWrap(this);"><br> -->
            <p><input type="button" class="butn" name="btnRecalcSubTotal" id="btnRecalcSubTotal" value="Recalculate subTotal" onclick="recalcSubTotal(this.form); return false;"><br>
            <input type="button" class="butn" name="btnRecalcTotal" id="btnRecalcTotal" value="Recalculate Totals" onclick="recalcOrderTotal(this); return false;"></p>
			<p><input class="butn" id="btnUpdateAlt" name="btnUpdateAlt" type="submit" value="Save Changes"></p>
          </td>
			<% If cblnShowBackOrderColumn Then %>
          <td width="40%" colspan="5" align=right style="PADDING-RIGHT: 12px;">
			<% Else %>
          <td width="40%" colspan="5" align=right style="PADDING-RIGHT: 12px;">
			<% End If %>
<%
Dim mcurDiscount

Dim mcurCoupon: mcurCoupon = 0
If cblnSF5AE Then
	If isNumeric(objRS.Fields("orderCouponDiscount").Value) Then mcurCoupon = CDbl(objRS.Fields("orderCouponDiscount").Value)
End If

mcurDiscount = Round(Abs(CDbl(objRS.Fields("orderGrandTotal").Value) - pcurRealSubTotal + mcurCoupon - CDbl(objRS.Fields("orderHandling").Value) - CDbl(objRS.Fields("orderCTax").Value) - CDbl(objRS.Fields("orderSTax").Value) - CDbl(objRS.Fields("orderShippingAmount").Value)), 2)
%>
<input type="hidden" name="baseOrderChanged" id="baseOrderChanged" value="">
<script language="javascript">

var newProductCounter = 0;
var GiftWrapSubTotal = 0;

function deleteProduct(theItem)
{

	if (! anyChecked(theDataForm.deleteodrdtID))
	{
		alert("Please select at least one product to delete from this order.");
		return false;
	}

	if (confirm("Are you sure you wish to delete the selected product(s)?"))
	{
		blnDeleteProduct = true;
		if (ValidInput(theDataForm)) {theDataForm.submit();}
	}
	
	return false;

}

var blnDeleteGiftWrap = false;
function deleteGiftWrap(theItem)
{

	if (! anyChecked(theDataForm.deleteodrdtAEID))
	{
		alert("Please select at least one gift wrap to remove from this order.");
		return false;
	}

	if (confirm("Are you sure you wish to remove the selected gift wrap(s)?"))
	{
		blnDeleteGiftWrap = true;
		if (ValidInput(theDataForm)){theDataForm.submit();}
	}
	
	return false;

}

function addProduct(theItem)
{
var pobjSourceRow = document.getElementById("sampleProductRow");
var pobjNewRow;
var pobjNewCell;

	pobjNewRow = document.all("tblOrderDetailSummary").insertRow(1);

	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input type='hidden' name='odrdtID' id='odrdtID." + newProductCounter + "' value='newProduct" + newProductCounter + "'>";

	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input name='odrdtProductID' id='odrdtProductID." + newProductCounter + "' tag='" + newProductCounter + "' style='text-align:left;' onchange='makeBaseOrderDirty();' value='' size='10'>"

	<% 
	If isArray(maryCustomDisplayColumns) Then
		For i = 0 To UBound(maryCustomDisplayColumns)
	%>
	pobjNewCell = pobjNewRow.insertCell();
	<%
		Next 'i
	End If
	%>
	
	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input name='odrdtProductName' id='odrdtProductName." + newProductCounter + "' tag='" + newProductCounter + "' style='text-align:left;' onchange='makeBaseOrderDirty();' value='' size='20' ondblclick=" + cstrQuote + "return false; openMovementWindow('FreeProductID','drdtProductName." + newProductCounter + "');" + cstrQuote + ">"

	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input name='odrdtQuantity' id='odrdtQuantity." + newProductCounter + "' tag='" + newProductCounter + "' style='text-align:center;' onchange='return recalcProductTotal(this);' value='' size='5'>"

	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input name='unitPrice' id='unitPrice." + newProductCounter + "' tag='" + newProductCounter + "' style='text-align:right;' onchange='return recalcProductTotal(this);' value='' size='5'>"

	pobjNewCell = pobjNewRow.insertCell();
	pobjNewCell.innerHTML = "<input name='odrdtSubTotal' id='odrdtSubTotal." + newProductCounter + "' tag='" + newProductCounter + "' style='text-align:right;' onchange='return recalcProductTotal(this);' value='' size='5'>"

	newProductCounter = newProductCounter ++;
}


function deleteOrderDetailAE(theItem)
{
var pobjItem = document.getElementById("odrdtGiftWrapPrice." + theItem.value);
pobjItem.value = 0;

recalcGiftWrapTotal(pobjItem);
recalcOrderTotal(theItem);

}

function deleteOrderDetail(theItem)
{
var pobjItem = document.getElementById("odrdtSubTotal." + theItem.value);
pobjItem.value = 0;

recalcSubTotal(theItem.form);
recalcOrderTotal(theItem);

}

function recalcGiftWrapTotal(theItem)
{
var theForm = theItem.form;
var pdblSubTotal = new Number(theForm.orderAmount.value);
var pstrTag = theItem.tag;
var pobjItem;
var pstrTemp;

pobjItem = document.getElementById("odrdtGiftWrapQTY." + pstrTag);
var podrdtQuantity = new Number(pobjItem.value);

pobjItem = document.getElementById("GiftWrapUnitPrice." + pstrTag);
var pdblunitPrice = new Number(pobjItem.value);

pobjItem = document.getElementById("odrdtGiftWrapPrice." + pstrTag);
var pdblodrdtSubTotal = new Number(pobjItem.value);

if (theItem.name == "odrdtGiftWrapPrice")
{
	pdblunitPrice = pdblodrdtSubTotal / podrdtQuantity;
	pobjItem = document.getElementById("GiftWrapUnitPrice." + pstrTag);
	pobjItem.value = pdblunitPrice;
}else{
	pdblodrdtSubTotal = pdblunitPrice * podrdtQuantity;
	pobjItem = document.getElementById("odrdtGiftWrapPrice." + pstrTag);
	pobjItem.value = pdblodrdtSubTotal;
}
recalcGiftWrapSubTotal(theItem.form);
recalcSubTotal(theItem.form);
recalcOrderTotal(theItem);

}

function recalcGiftWrapSubTotal(theForm)
{

	if (theForm.odrdtGiftWrapPrice == null){return false;}
	
var plngCount = theForm.odrdtGiftWrapPrice.length;
var i;
var pdblSubTotal = 0;
var pdblExtPrice;

	if (plngCount == undefined)
	{
		GiftWrapSubTotal = new Number(theForm.odrdtGiftWrapPrice.value);
	}else{
		for (i=0; i < plngCount;i++)
		{
			pdblExtPrice = new Number(theForm.odrdtGiftWrapPrice[i].value);
			pdblSubTotal = pdblSubTotal + pdblExtPrice;
		}
	}
	GiftWrapSubTotal = pdblSubTotal;
}

function recalcProductTotal(theItem)
{
var theForm = theItem.form;
var pdblSubTotal = new Number(theForm.orderAmount.value);
var pstrTag = theItem.tag;
var pobjItem;
var pstrTemp;

pobjItem = document.getElementById("odrdtQuantity." + pstrTag);
var podrdtQuantity = new Number(pobjItem.value);

pobjItem = document.getElementById("unitPrice." + pstrTag);
var pdblunitPrice = new Number(pobjItem.value);

pobjItem = document.getElementById("odrdtSubTotal." + pstrTag);
var pdblodrdtSubTotal = new Number(pobjItem.value);

if (theItem.name == "odrdtSubTotal")
{
	pdblunitPrice = pdblodrdtSubTotal / podrdtQuantity;
	pobjItem = document.getElementById("unitPrice." + pstrTag);
	pobjItem.value = pdblunitPrice;
}else{
	pdblodrdtSubTotal = pdblunitPrice * podrdtQuantity;
	pobjItem = document.getElementById("odrdtSubTotal." + pstrTag);
	pobjItem.value = pdblodrdtSubTotal;
}

recalcSubTotal(theItem.form);
recalcOrderTotal(theItem);

}

function recalcSubTotal(theForm)
{
var plngCount = theForm.odrdtSubTotal.length;
var i;
var pdblSubTotal = 0;
var pdblExtPrice;

	recalcGiftWrapSubTotal(theForm);
	if (plngCount == undefined)
	{
		pdblSubTotal = new Number(theForm.odrdtSubTotal.value);
	}else{
		for (i=0; i < plngCount;i++)
		{
			pdblExtPrice = new Number(theForm.odrdtSubTotal[i].value);
			pdblSubTotal = pdblSubTotal + pdblExtPrice;
		}
	}
	theForm.orderAmount.value = pdblSubTotal + GiftWrapSubTotal;
}

function recalcOrderTotal(theItem)
{
var theForm = theItem.form;
var pdblSubTotal = new Number(theForm.orderAmount.value);
var pdblorderCouponDiscount = new Number(theForm.orderCouponDiscount.value);
var pdblDiscount = new Number(theForm.Discount.value);
var pdblorderSTax = new Number(theForm.orderSTax.value);
var pdblorderCTax = new Number(theForm.orderCTax.value);
var pdblorderShippingAmount = new Number(theForm.orderShippingAmount.value);
var pdblorderHandling = new Number(theForm.orderHandling.value);

var pdblorderGrandTotal = pdblSubTotal - pdblorderCouponDiscount - pdblDiscount + pdblorderSTax + pdblorderCTax + pdblorderShippingAmount + pdblorderHandling;
pdblorderGrandTotal = pdblorderGrandTotal.toFixed(2);
theForm.orderGrandTotal.value = pdblorderGrandTotal;

makeBaseOrderDirty();
}

function makeBaseOrderDirty()
{
document.all("baseOrderChanged").value = 1;
}

</script>

            <table border="0" cellspacing="0" cellpadding="0">
			  <colgroup align=left>
			  <colgroup align=right>
              <tr>
                <td>Subtotal:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                <td><input name="orderAmount" id="orderAmount" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= pcurRealSubTotal %>" size="5"></td>
              </tr>
			  <% If cblnSF5AE Then %>
              <tr>
                <td>Coupon (<input name="orderCouponCode" id="orderCouponCode" style="text-align:center;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderCouponCode").Value & "") %>" size="5">):</td>
                <td><input name="orderCouponDiscount" id="orderCouponDiscount" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderCouponDiscount").Value & "") %>" size="5"></td>
              </tr>
			  <% Else %>
			  <input type="hidden" name="orderCouponDiscount" id="Hidden1" value="0">
			  <% End If %>
              <tr>
                <td>Discount:</td>
                <td><input name="Discount" id="Discount" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= mcurDiscount %>" size="5"></td>
              </tr>
              <tr>
                <td><%= objRS.Fields("cshpaddrShipState").Value %> Tax:</td>
                <td><input name="orderSTax" id="orderSTax" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderSTax").Value & "") %>" size="5"></td>
              </tr>
              <tr>
                <td><%= objRS.Fields("cshpaddrShipCountry").Value %> Tax:</td>
                <td><input name="orderCTax" id="orderCTax" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderCTax").Value & "") %>" size="5"></td>
              </tr>
              <tr>
                <td>Shipping (<input name="orderShipMethod" id="orderShipMethod" style="text-align:center;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderShipMethod").Value & "") %>" size="<%= Len(Trim(objRS.Fields("orderShipMethod").Value & "")) + 2 %>">):&nbsp;&nbsp;</td>
                <td><input name="orderShippingAmount" id="orderShippingAmount" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderShippingAmount").Value & "") %>" size="5"></td>
              </tr>
              <tr>
                <td>Handling:</td>
                <td><input name="orderHandling" id="orderHandling" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderHandling").Value & "") %>" size="5"></td>
              </tr>
              <tr>
                <td colspan="2">
                  <hr>
                </td>
              </tr>
              <tr>
                <td>Total:</td>
                <td><input name="orderGrandTotal" id="orderGrandTotal" style="text-align:right;" onchange="return recalcOrderTotal(this);" value="<%= Trim(objRS.Fields("orderGrandTotal").Value & "") %>" size="5"></td>
              </tr>
            <% 
            If cblnSF5AE Then
				If objRS.Fields("orderBackOrderAmount").Value > 0 Or cblnShowBilledAmountRow Then
            %>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
				<tr>
					<td align="left">Billed Amount:</td>
					<td align="right"><input name="orderBillAmount" id="orderBillAmount" style="text-align:right;" onchange="recalcOrderTotal(this); this.form.orderBackOrderAmount.value = this.form.orderGrandTotal.value - this.value;" value="<%= objRS.Fields("orderBillAmount").Value %>" size="5"></td>
				</tr>
				<tr>
					<td align="left">Amount Remaining:</td>
					<td align="right"><input name="orderBackOrderAmount" id="Text1" style="text-align:right;" onchange="recalcOrderTotal(this); this.form.orderBillAmount.value = this.form.orderGrandTotal.value - this.value;" value="<%= objRS.Fields("orderBackOrderAmount").Value %>" size="5"></td>
				</tr>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
            <% 
				End If	' 
            End If	'cblnSF5AE 
            %>
            <% 
            If cblnAddon_GCMgr Then 
				If GC_LoadByOrder(objRS.Fields("OrderID").Value, objRS.Fields("orderGrandTotal").Value) Then
            %>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
				<tr>
					<td align="left">Certificate (<a href="ssGiftCertificateAdmin.asp?Action=ViewByCode&ssGCCode=<%= mstrCertificate %>"><%= mstrCertificate %></a>):</td>
					<td align="right"><%= FormatCurrency(mdblssCertificateAmount, 2) %></td>
				</tr>
				<tr>
					<td align="left">Amount Billed:</td>
					<td align="right"><%= FormatCurrency(mdblssGCNewTotalDue, 2) %></td>
				</tr>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
            <% 
				End If	'GC_LoadByOrder 
            End If	'cblnAddon_GCMgr 
            %>
            </table>
          </td>
        </tr>
		<% If Len(objRS.Fields("orderComments").Value & "") > 0 Then %>
        <tr>
          <td colspan=5 align=left>&nbsp;&nbsp;Special Instructions:&nbsp;&nbsp;<%= objRS.Fields("orderComments").Value %></td>
        </tr>
        <% End If %>
      </table>
    </TD>
  </TR>
  <tr class="tblhdr">
    <th colspan="2">Customer Information</th>
  </tr>
  <tr>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%" ID="Table4">
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td><a href="" title="Click to Edit Billing Information" onclick="OpenHelp('ssOrderAdmin_Customer.asp?Action=EditBilling&ID=<%= objRS.Fields("custID").Value %>'); return false;">Sold to:</a>&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(objRS.Fields("custCompany").Value & "") > 0 Then %>
        <tr>
          <td>Company:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>Address:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custAddr1").Value %></td>
        </tr>
		<% If Len(objRS.Fields("custAddr2").Value & "") > 0 Then %>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("custAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>
          <%
          If mblnShowFullCountryName Then
			Response.Write objRS.Fields("billToCountryName").Value
		  Else
			Response.Write objRS.Fields("custCountry").Value
		  End If
          %>
          </td>
        </tr>
		<% If Len(objRS.Fields("custPhone").Value & "") > 0 Then %>
        <tr>
          <td>Phone:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(objRS.Fields("custFax").Value & "") > 0 Then %>
        <tr>
          <td>Fax:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td valign="top"><a title="Send email" href="mailto:<%= objRS.Fields("custEmail").Value %>&subject=Order <%= objRS.Fields("OrderID").Value %>&body=">email:</a>&nbsp;&nbsp;</td>
          <td valign="top"><%= objRS.Fields("custEmail").Value %><br>PW= <%= objRS.Fields("custPasswd").Value %></td>
        </tr>
      </table>
    </td>
<%
pstrCustName = Replace(objRS.Fields("cshpaddrShipFirstName").Value & " " & objRS.Fields("cshpaddrShipMiddleInitial").Value & " " & objRS.Fields("cshpaddrShipLastName").Value,"  "," ")
pstrCustAddr = objRS.Fields("cshpaddrShipCity").Value & ", " & objRS.Fields("cshpaddrShipState").Value & " " & objRS.Fields("cshpaddrShipZip").Value

Dim pstrShippingAddressColor
If isAddressDifferent(objRS) Then
	pstrShippingAddressColor = " bgcolor=yellow"
Else
	pstrShippingAddressColor = ""
End If
%>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%" <%= pstrShippingAddressColor %>>
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td><a href="" title="Click to Edit Shipping Information" onclick="OpenHelp('ssOrderAdmin_Customer.asp?Action=EditShipping&ID=<%= objRS.Fields("cshpaddrID").Value %>'); return false;">Shipped to:</a>&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipCompany").Value & "") > 0 Then %>
        <tr>
          <td>Company:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>Address:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipAddr1").Value %></td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipAddr2").Value & "") > 0 Then %>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>
          <%
          If mblnShowFullCountryName Then
			Response.Write objRS.Fields("shipToCountryName").Value
		  Else
			Response.Write objRS.Fields("cshpaddrShipCountry").Value
		  End If
          %>
          </td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipPhone").Value & "") > 0 Then %>
        <tr>
          <td>Phone:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(objRS.Fields("cshpaddrShipFax").Value & "") > 0 Then %>
        <tr>
          <td>Fax:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td><a title="Send email" href="mailto:<%= objRS.Fields("cshpaddrShipEmail").Value %>&subject=Order <%= objRS.Fields("OrderID").Value %>&body=">email:</a>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipEmail").Value %></td>
        </tr>
      </table>
    </td>
  </tr>
<% If mclsOrder.PriorShippedOrders > 0 Then %>
<tr><td colspan=2><hr><strong>Prior Shipped Order Count: <%= mclsOrder.PriorShippedOrders %></strong><hr></td></tr>
<% End If %>

   <tr class="tblhdr">
    <th width="100%" colspan="2">Payment Method</th>
   </tr>
   <tr>
     <td valign=top style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
		<%
		Dim pblnPhoneFax
		Dim pblnPayPal
		Dim mstrOrderPaymentMethod
		Dim mstrOrderPaymentMethodText
		Dim mstrOrderPaymentMethodAutoFill

		mstrOrderPaymentMethod = Trim(objRS.Fields("orderPaymentMethod").Value & "")
		
		'Now PayPal may or may not just be PayPal
		pblnPayPal = CBool(Instr(1, mstrOrderPaymentMethod, "PayPal") > 0)
		
		'Now PhoneFax will/may have a payment type extension. Ex. PhoneFax_Credit Card, PhoneFax_eCheck, PhoneFax_PO
		pblnPhoneFax = CBool(Instr(1, mstrOrderPaymentMethod, "PhoneFax") > 0)
		mstrOrderPaymentMethod = Replace(mstrOrderPaymentMethod, "PhoneFax_", "")	'Now remove the prefix

		Select Case mstrOrderPaymentMethod
			Case "Credit Card"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
			Case "eCheck"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod & objRS.Fields("orderCheckNumber").Value
								mstrOrderPaymentMethodAutoFill = "eCheck " & objRS.Fields("orderCheckNumber").Value
			Case "PO"
								mstrOrderPaymentMethodText = "P.O. " & objRS.Fields("orderPurchaseOrderNumber").Value & "<BR>" _
														   & "Name: " & objRS.Fields("orderPurchaseOrderName").Value
								mstrOrderPaymentMethodAutoFill = "P.O. " & objRS.Fields("orderPurchaseOrderNumber").Value
			Case "PayPal"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
			Case Else
				'Catch the unique ones here
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod				
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
		End Select
		If pblnPhoneFax Then mstrOrderPaymentMethodText = "Phone/Fax - " & mstrOrderPaymentMethodText
		
		%>
		<table class="tbl" width="100%" cellpadding="0" cellspacing="0" border="0" rules="none">
		  <tr><th align="left"><%= mstrOrderPaymentMethodText %></th></tr>
		  <tr><td>
				<input type="hidden" name="origPaymentType" id="origPaymentType" value="<%= mstrOrderPaymentMethod %>">
		  		<select name="newPaymentType" id="newPaymentType"><%= createCombo("SELECT Distinct orderPaymentMethod FROM sfOrders ORDER BY orderPaymentMethod", "", "orderPaymentMethod", mstrOrderPaymentMethod) %></select>
		  </td></tr>

		  <%
			Select Case mstrOrderPaymentMethod
				Case "Credit Card":
					Response.Write "<tr><td>"
					Call ShowCC(plngOrderID)
					Response.Write "</td></tr>"
				Case "eCheck":
				Case "PO"
				Case "PayPal"
				Case Else
					'Catch the unique ones here
			End Select
		  %>
	    </table>
     </td>
     <td valign=top align="left" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
		<%
			Select Case mstrOrderPaymentMethod
				Case "Credit Card":		Call ShowTransactionResponse(plngOrderID)
				Case "eCheck":			Response.Write "<p align=center><a href='../printEcheck.asp?orderid=" & plngOrderID & "'>Print Check</a></p>"
				Case "PO"
				Case "PayPal"
				Case Else
					'Catch the unique ones here
			End Select
		%>
	  </td>
	</tr>
<script language="javascript">
function SetPaymentMethod()
{
document.frmData.ssPaidVia.value = "<%= mstrOrderPaymentMethodAutoFill %>";
}
</script>

   <tr class="tblhdr">
    <th colspan="2">Referral Information</th>
   </tr>
   <tr>
    <th colspan="2"align="left" >
	  <table cellpadding="0" cellspacing="0">
		<tr>
		  <td align="right">Trading Partner:</td>
		  <td>&nbsp;<%= objRS.Fields("orderTradingPartner").Value %></td>
		</tr>
		<tr>
		  <td align="right">Remote Address:</td>
		  <td>&nbsp;<%= Trim(objRS.Fields("orderRemoteAddress").Value & "") %>
		  &nbsp;<a href="http://www.whois.sc/<%= Trim(objRS.Fields("orderRemoteAddress").Value & "") %>" target="whois">(Whois Lookup)</a>
		  &nbsp;<a href="http://www.dnsstuff.com/tools/ptr.ch?ip=<%= Trim(objRS.Fields("orderRemoteAddress").Value & "") %>" target="ReverseDNS">(Reverse DNS Lookup)</a>
		  &nbsp;<a href="http://www.dnsstuff.com/tools/netgeo.ch?ip=<%= Trim(objRS.Fields("orderRemoteAddress").Value & "") %>" target="NetGeo">(NetGeo Lookup)</a>
		  </td>
		</tr>
		<tr>
		  <td valign="top"  align="right">Http Referer:</td>
		  <td>&nbsp;<%= objRS.Fields("orderHttpReferrer").Value %></td>
		</tr>
	  </table>
    </th>
   </tr>

 <tr><td colspan="2" height="3pt"></td></tr>
 <tr>
    <td width="100%" colspan="2">
	  <table id="tblPricing" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="0" cellspacing="0" border="0" rules="none">
	    <tr>
		  <th id="tdStatus" class="hdrSelected" nowrap onclick="return DisplaySection('Status');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Order Status">&nbsp;Order&nbsp;Status&nbsp;</th>
		  <th nowrap width="2pt"></th>
		  <th id="tdBackOrder" class="hdrNonSelected" nowrap onclick="return DisplaySection('BackOrder');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Back order information">&nbsp;Back Order Information&nbsp;</th>
		  <th nowrap width="2pt"></th>
		  <th id="tdExportStatus" class="hdrSelected" nowrap onclick="return DisplaySection('ExportStatus');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Export Status">&nbsp;Export Status&nbsp;</th>
<% If isArray(mclsOrder.CustomOrderValues) Then %>
		  <th nowrap width="2pt"></th>
		  <th id="tdCustom" class="hdrSelected" nowrap onclick="return DisplaySection('Custom');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Custom values">&nbsp;Custom&nbsp;</th>
<%
	plngNumberDetailColumns = 8
   Else
	plngNumberDetailColumns = 6
   End If	'cblnUseCustomOrderFields %>
		  <th width="90%">&nbsp;</th>
	    </tr>
		<tr>
			<td colspan="<%= plngNumberDetailColumns %>" class="hdrSelected" height="1px"></td>
		</tr>
	    <tr>
	    <td colspan="<%= plngNumberDetailColumns %>" align="left" valign="top">
<!-- tblStatus -->
		<table id="tblStatus" class="tbl" cellpadding="3" cellspacing="0" border="0" rules="none">
        <tr>
          <td>Order <%= mclsOrder.ssOrderID %> Placed On:&nbsp;&nbsp;</td>
          <td><%= FormatDateTime(objRS.Fields("orderDate").Value,1) & " at " & FormatDateTime(objRS.Fields("orderDate").Value,3) %></td>
        </tr>
		<tr>
			<td align="right">&nbsp;</td>
			<td align="left"><input type="checkbox" id="orderIsComplete" name="orderIsComplete" value='1' <%= isChecked(objRS.Fields("orderIsComplete").Value) %>><label for="orderIsComplete">Order Set Complete</label>&nbsp;</td>
		</tr>
		<tr>
			<td align="right">&nbsp;</td>
			<td align="left"><input type="checkbox" id="orderVoided" name="orderVoided" value='1' <%= isChecked(objRS.Fields("orderVoided").Value) %>><label for="orderVoided">Order Voided</label>&nbsp;</td>
		</tr>
		<tr>
			<td align="right">&nbsp;</td>
			<td align="left"><input type="checkbox" id=ssOrderFlagged name=ssOrderFlagged value='1' <% If objRS.Fields("ssOrderFlagged").Value Then Response.Write "checked"%>><label for="ssOrderFlagged">Flag this order</label>&nbsp;</td>
		</tr>
        <tr>
          <td><label for="ssOrderStatus" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">External Order Status:</label>&nbsp;&nbsp;</td>
          <td>
        	<select name="ssOrderStatus" id="ssOrderStatus">
			<% For i = 0 To UBound(maryOrderStatuses) %>
			<% If i = objRS.Fields("ssOrderStatus").Value Then %>
			<option value="<%= i %>" selected><%= maryOrderStatuses(i) %></option>
			<% Else %>
			<option value="<%= i %>"><%= maryOrderStatuses(i) %></option>
			<% End If %>
			<% Next 'i %>
			</select>
          </td>
        </tr>
        <tr>
          <td><label for="ssInternalOrderStatus" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Internal Order Status:</label>&nbsp;&nbsp;</td>
          <td>
        	<select name="ssInternalOrderStatus" id="ssInternalOrderStatus">
			<% For i = 0 To UBound(maryInternalOrderStatuses) %>
			<% If i = objRS.Fields("ssInternalOrderStatus").Value Then %>
			<option value="<%= i %>" selected><%= maryInternalOrderStatuses(i)(0) %></option>
			<% Else %>
			<option value="<%= i %>"><%= maryInternalOrderStatuses(i)(0) %></option>
			<% End If %>
			<% Next 'i %>
			</select>
          </td>
        </tr>
        <tr>
          <td><label for="ssDatePaymentReceived" onmouseover="showDataEntryTip(this);" onMouseOut="htm();" ondblclick="document.frmData.ssDatePaymentReceived.value='<%= Date() %>'">Date Payment Received:</label>&nbsp;&nbsp;</td>
          <td>
            <input id=ssDatePaymentReceived name=ssDatePaymentReceived Value="<%= objRS.Fields("ssDatePaymentReceived").Value %>">
			<a HREF="javascript:doNothing()" title="Select start date"
			onClick="setDateField(document.frmData.ssDatePaymentReceived); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<img SRC="images/calendar.gif" BORDER=0></a>
          </td>
        </tr>
        <tr>
          <td><label for="ssPaidVia" ondblclick="SetPaymentMethod();" title="Double click to set payment method">Paid By:</label>&nbsp;&nbsp;</td>
          <td><input id=ssPaidVia name=ssPaidVia Value="<%= objRS.Fields("ssPaidVia").Value %>"></td>
        </tr>

        <tr>
          <td><label for="ssDateOrderShipped" ondblclick="setShipDate();" title="Double click to set date order shipped to today's date">Date Order Shipped:</label>&nbsp;&nbsp;</td>
          <td>
            <input id=ssDateOrderShipped name=ssDateOrderShipped Value="<%= objRS.Fields("ssDateOrderShipped").Value %>">
			<a HREF="javascript:doNothing()" title="Select start date"
			onClick="setDateField(document.frmData.ssDateOrderShipped); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<img SRC="images/calendar.gif" BORDER=0></a>&nbsp;
			<% If Len(cstrSpecial) > 0 Then %><a href="" title="Send confirmation email" onclick="DoSpecial(<%= mclsOrder.ssOrderID %>); return false;"><%= cstrSpecial %></a><% End If %>
          </td>
        </tr>
        <tr>
          <td><label for="ssShippedVia">Order Shipped via:&nbsp;&nbsp;</label></td>
          <td>
            <select id="ssShippedVia" name="ssShippedVia">
              <%
			  If Len(objRS.Fields("ssShippedVia").Value & "") = 0 Then
				Response.Write "<option value=''></option>"
			  End If
              For i=0 to UBound(maryShipMethods)
				If CStr(objRS.Fields("ssShippedVia").Value & "") = CStr(i) Then
					Response.Write "<option value='" & i & "' selected>" & maryShipMethods(i)(0) & "</option>"
				Else
					Response.Write "<option value='" & i & "'>" & maryShipMethods(i)(0) & "</option>"
				End If
              Next
			  %>
            </select>
          </td>
        </tr>
        <tr>
          <td><label for="ssTrackingNumber">Tracking Number:</label>&nbsp;&nbsp;</td>
          <td><textarea id=ssTrackingNumber name=ssTrackingNumber rows="5" cols="50" title="Separate tracking numbers from carriers with semi-colons (;). Each entry should be on a separate line."><%= Trim(objRS.Fields("ssTrackingNumber").Value) %></textarea></td>
        </tr>
        <tr>
          <td><label for="ssInternalNotes">Internal Notes:</label>&nbsp;&nbsp;</td>
          <td>
            <textarea id=ssInternalNotes name=ssInternalNotes rows="5" cols="50"><%= objRS.Fields("ssInternalNotes").Value %></textarea>
            <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.ssInternalNotes);" title="Edit this field with the HTML Editor"><img SRC="images/properites.gif" BORDER=0></a>
          </td>
        </tr>
        <tr>
          <td><label for="ssExternalNotes">External Notes:</label>&nbsp;&nbsp;</td>
          <td>
            <textarea id=ssExternalNotes name=ssExternalNotes rows="5" cols="50"><%= objRS.Fields("ssExternalNotes").Value %></textarea>
            <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.ssExternalNotes);" title="Edit this field with the HTML Editor"><img SRC="images/properites.gif" BORDER=0></a>
          </td>
        </tr>
	<%
	Dim paryCustomValues
	paryCustomValues = mclsOrder.CustomValues
	If isArray(paryCustomValues) Then
		For i = 0 To UBound(paryCustomValues)
	%>
      <tr>
        <td class="Label"><%= paryCustomValues(i)(0) %>:</td>
        <td>
        <%= writeHTMLFormElement(paryCustomValues(i)(3), paryCustomValues(i)(4), paryCustomValues(i)(1), paryCustomValues(i)(1), paryCustomValues(i)(2), paryCustomValues(i)(5), "MakeDirty(this);") %>
        </td>
      </tr>
    <%
		Next 'i
	End If
	%>
	</table>
<!-- tblStatus -->

<table id="tblBackOrder" class="tbl" style="display:none;" cellpadding="3" cellspacing="0" border="0" rules="none">
      <tr>
        <td>
		  <% If cblnUseBackOrder Then %>
			<tr>
			  <td align="right"><label for="ssBackOrderDateNotified" ondblclick="document.frmData.ssBackOrderDateNotified.value='<%= Date() %>'" title="Double click to set to today's date">Date Back Order Notification Sent:</label>&nbsp;&nbsp;</td>
			  <td align="left">
			    <input id=ssBackOrderDateNotified name=ssBackOrderDateNotified Value="<%= objRS.Fields("ssBackOrderDateNotified").Value %>">
				<a HREF="javascript:doNothing()" title="Select start date"
				onClick="setDateField(document.frmData.ssBackOrderDateNotified); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
				<img SRC="images/calendar.gif" BORDER=0></a>
			  </td>
			</tr>
			<tr>
			  <td align="right"><label for="ssBackOrderDateExpected" ondblclick="document.frmData.ssBackOrderDateExpected.value='<%= Date() %>'" title="Double click to set to today's date">Date Expected In:</label>&nbsp;&nbsp;</td>
			  <td align="left">
			    <input id=ssBackOrderDateExpected name=ssBackOrderDateExpected Value="<%= objRS.Fields("ssBackOrderDateExpected").Value %>">
				<a HREF="javascript:doNothing()" title="Select start date"
				onClick="setDateField(document.frmData.ssBackOrderDateExpected); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
				<img SRC="images/calendar.gif" BORDER=0></a>
			  </td>
			</tr>
			<tr>
			  <td align="right"><label for="ssBackOrderTrackingNumber">Back Order Tracking Numbers:</label>&nbsp;&nbsp;</td>
			  <td align="left">
			    <textarea name=ssBackOrderTrackingNumber id="ssBackOrderTrackingNumber" rows="5" cols="50"><%= Trim(objRS.Fields("ssBackOrderTrackingNumber").Value) %></textarea>
			  </td>
			</tr>
			<tr>
			  <td align="right"><label for="ssBackOrderMessage">Back Order Message:</label>&nbsp;&nbsp;</td>
			  <td align="left">
			    <textarea id=ssBackOrderMessage name=ssBackOrderMessage rows="5" cols="50"><%= objRS.Fields("ssBackOrderMessage").Value %></textarea>
                <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.ssBackOrderMessage);" title="Edit this field with the HTML Editor"><img SRC="images/properites.gif" BORDER=0></a>
			  </td>
			</tr>
			<tr>
			  <td align="right"><label for="ssBackOrderInternalMessage">Internal Back Order Message:</label>&nbsp;&nbsp;</td>
			  <td align="left">
			    <textarea id=ssBackOrderInternalMessage name=ssBackOrderInternalMessage rows="5" cols="50"><%= objRS.Fields("ssBackOrderInternalMessage").Value %></textarea>
                <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.ssBackOrderInternalMessage);" title="Edit this field with the HTML Editor"><img SRC="images/properites.gif" BORDER=0></a>
			  </td>
			</tr>
		  <% End If	'cblnUseBackOrder %>
        </td>
      </tr>
</table>
<!-- tblStatus -->

<%
paryCustomValues = mclsOrder.CustomOrderValues
If isArray(paryCustomValues) Then
%>
<table id="tblCustom" class="tbl" width="100%" style="display:none;" cellpadding="3" cellspacing="0" border="0" rules="none">
    <% For i = 0 To UBound(paryCustomValues) %>
      <tr>
        <td class="Label"><%= paryCustomValues(i)(0) %>:</td>
        <td>
        <%= writeHTMLFormElement(paryCustomValues(i)(3), paryCustomValues(i)(4), paryCustomValues(i)(1), paryCustomValues(i)(1), paryCustomValues(i)(2), paryCustomValues(i)(5), "MakeDirty(this);") %>
        </td>
      </tr>
    <% Next 'i %>
</table>
<% End If	'isArray(paryCustomValues) %>

<table id="tblExportStatus" class="tbl" style="display:none;" cellpadding="3" cellspacing="0" border="0" rules="none">
	<tr>
		<td align="right"><label for="ssExportedPayment">Exported to Payment Processor</label>&nbsp;</td>
		<td align="left"><input type="checkbox" id="ssExportedPayment" name="ssExportedPayment" value='1' <% If objRS.Fields("ssExportedPayment").Value Then Response.Write "checked"%>></td>
	</tr>
	<tr>
		<td align="right"><label for="ssExportedShipping">Exported to Shipping Program</label>&nbsp;</td>
		<td align="left"><input type="checkbox" id="ssExportedShipping" name="ssExportedShipping" value='1' <% If objRS.Fields("ssExportedShipping").Value Then Response.Write "checked"%>></td>
	</tr>
	<tr>
		<td align="right"><label for="ssExported">Exported to Accounting Program</label>&nbsp;</td>
		<td align="left"><input type="checkbox" id="ssExported" name="ssExported" value='1' <% If objRS.Fields("ssExported").Value Then Response.Write "checked"%>></td>
	</tr>
</table>

	    </td>
	    </tr>
	  </table>
    </td>
  </tr>
 <tr class="tblhdr"><td colspan="2" height="3pt">Email</td></tr>
  <tr>
    <td colspan="2">

      <table border="0" cellspacing="0" ID="tblOrderStatusDetail">
		<colgroup align=right>
		<colgroup align=left>
		<tr>
		</tr>
		
<%
Dim mstremailFile
Dim pstrEmailBody
Dim pstrEmailSubject

Call LoadEmails(mstremailFile, pstrEmailSubject, pstrEmailBody, objRS, False)

%>
 <tr><td colspan=2 align=left>
	<div id="divEmail" style="position:absolute; display:none">
<table border="3" cellspacing="0" cellpadding="0" bgcolor="white" id="tblEmail" style="border-style:outset; border-color:steelblue;"><tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="3" ID="Table8">
  <tr>
    <td align="right">
      <p>Select an email template:</td>
    <td>
      <script language="javascript">
      function changeEmailTemplate(theSelect)
      {
      theSelect.form.emailSubject.value  = document.all("enEmail_Subject" + theSelect.selectedIndex).value;
      
      theSelect.form.emailBody.value  = document.all("enEmail_Body" + theSelect.selectedIndex).value;
      }
      </script>
      <% For i = 0 To UBound(maryEmails) %>
      <input type="hidden" name="enEmail_Subject<%= i %>" id="enEmail_Subject<%= i %>" value="<%= maryEmails(i)(enEmail_Subject) %>">
      <input type="hidden" name="enEmail_Body<%= i %>" id="enEmail_Body<%= i %>" value="<%= maryEmails(i)(enEmail_Body) %>">
      <% Next 'i %>
      <select name="emailFile" ID="emailFile" onchange="changeEmailTemplate(this); return false;">
      <% For i = 0 To UBound(maryEmails) %>
      <%   If CBool(mstremailFile = maryEmails(i)(enEmail_FileName)) Or CBool(Len(mstremailFile)=0 And (i = UBound(maryEmails))) Then  %>
      <%	pstrEmailSubject = maryEmails(i)(enEmail_Subject) %>
      <%	pstrEmailBody = maryEmails(i)(enEmail_Body) %>
		<option selected><%= maryEmails(i)(enEmail_FileName) %></option>
      <%   Else %>
		<option><%= maryEmails(i)(enEmail_FileName) %></option>
      <%   End If %>
      <% Next 'i %>
      </select>
    </td>
  </tr>
  <tr>
    <td align="right">
      <p>To:</td>
    <td><input type="text" name="emailTo" ID="emailTo" size="75" VALUE="<%= objRS.Fields("custEmail").Value %>"></td>
  </tr>
  <tr>
    <td align="right">Subject:</td>
    <td><input type="text" name="emailSubject" ID="emailSubject" size="75" VALUE="<%= pstrEmailSubject %>"></td>
  </tr>
  <tr>
    <td align="right">Body:</td>
    <td><textarea rows="9" name="emailBody" ID="emailBody" cols="70"><%= pstrEmailBody %></textarea></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input class='butn' type="button" value="Send" name="btnSendEmail" ID="btnSendEmail" onclick="document.frmData.SendEmail.checked=true; document.all('divEmail').style.display = 'none';">&nbsp;
        <input class='butn' type="button" value="Cancel" name="Cancel" ID="Cancel" onclick='document.all("divEmail").style.display = "none";'></td>
		<input type=hidden id=StockEmail name=StockEmail value=1>
  </tr>
</table>
</td></tr></table>
</div>
</td></tr>

<% 'End If %>		

<% If Len(objRS.Fields("ssDateEmailSent").Value & "") > 0 Then %>
    <tr>
        <td colspan="2" align="left">Order Fulfillment Email Sent on:&nbsp;&nbsp;<%= FormatDateTime(objRS.Fields("ssDateEmailSent").Value,1) %></td>
    </tr>
<% End If %>
        <tr>
          <td colspan="2" align="left"><input type="checkbox" id="Checkbox2" name="SendEmail" value="1">&nbsp;<label for"SendEmail">Send Email</label>&nbsp;
		    <a href="" onclick='document.all("divEmail").style.display = ""; ReplaceEmailText(); document.frmData.StockEmail.value=0; document.frmData.emailBody.focus(); ScrollToElem("btnSendEmail"); return false;'>(Edit)</a>
          </td>
        </tr>
 <tr><td><hr /></td></tr>
        <tr>
          <td colspan="2" align=left>
            <font size="-1">
			<a href="" title="Open order detail" onclick="OpenHelp('ssOrderAdmin_PrintableDetail.asp?Action=ViewOrder&OrderID=<%= mclsOrder.ssOrderID %>'); return false;">Printable Version</a><br>
			<!--
			<a href="" title="Send confirmation email" onclick="SendConfirmationEmail(<%= mclsOrder.ssOrderID %>); return false;">Send Confirmation Email</a><br>
			-->
            <a href="<% Response.Write cstrWebSite & "?OrderID=" & objRS.Fields("OrderID").Value & "&email=" & Trim(objRS.Fields("custEmail").Value) %>">View Customer's Order Status Page</a><br>
            <a href="<% Response.Write cstrWebSite & "?OrderID=" & objRS.Fields("OrderID").Value & "&email=" & Trim(objRS.Fields("custEmail").Value) & "&Password=" & Trim(objRS.Fields("custPasswd").Value) %>">Impersonate Customer</a><br>
            
            <% If cblnUseASPPages Then %>
			<a href="ssOrderAdmin_Invoice.asp?Action=ViewOrder&OrderID=<%= mclsOrder.ssOrderID %>" title="Open invoice" target="Invoice">View Invoice</a><br>
			<a href="ssOrderAdmin_PackingSlip.asp?Action=ViewOrder&OrderID=<%= mclsOrder.ssOrderID %>" title="Open packing slip" target="PackingSlip">View Packing Slip</a>
            <% Else %>
			<a href="" onclick="exportTemplates_Alternate('<%= cstrInvoiceTemplate %>', <%= mclsOrder.ssOrderID %>, 'Invoice'); return false;" title="Open invoice">View Invoice</a><br>
			<a href="" onclick="exportTemplates_Alternate('<%= cstrPackingSlipTemplate %>', <%= mclsOrder.ssOrderID %>, 'PackingSlip'); return false;" title="Open packing slip">View Packing Slip</a>
			<% End If 'cblnUseASPPages %>
			</font>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <%
	'Call ShowCustomPayPalActions
	'Call ShowCustomPayPalActions
  %>
</TABLE>
<script language=javascript>
function exportTemplates_Alternate(strTemplate, lngOrderID, strTarget)
{

	theDataForm.action='ssOrderAdmin_Export.asp';
	theDataForm.Action.value = 'viewOrders' + '|' + strTemplate;
	theDataForm.target=strTarget;
	theDataForm.submit();
	theDataForm.action='ssOrderAdmin.asp';
	theDataForm.target='';
	return false;
}
</script>
<% 

End Sub 'ShowOrderDetail

'**************************************************************************************************************************************************

Sub ShowTransactionResponse(lngOrderID)

Dim pstrSQL
Dim pobjRSCC

	pstrSQL = "SELECT trnsrspID, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess FROM sfTransactionResponse WHERE trnsrspOrderId=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
	Set pobjRSCC = GetRS(pstrSQL)
	If Not pobjRSCC.EOF Then
		Call ShowTransactionResponse_Complete(pobjRSCC)
		'Call ShowTransactionResponse_short(pobjRSCC)
	End If	'pobjRSCC.EOF
	
	Call ReleaseObject(pobjRSCC)

End Sub	'ShowTransactionResponse

'**************************************************************************************************************************************************

Sub ShowTransactionResponse_Complete(byRef objRSCC)

Dim paryCodes
Dim pstrAVSCode
Dim pstrCCVCode

	If objRSCC.RecordCount >= 1 Then
		Response.Write "<table class=""tbl"" width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" rules=""none"">"
		For i = 1 To objRSCC.RecordCount
			paryCodes = Split(objRSCC.Fields("trnsrspAVSCode").Value & "", "|")
			If UBound(paryCodes) >= 0 Then pstrAVSCode = paryCodes(0)
			If UBound(paryCodes) >= 1 Then pstrCCVCode = paryCodes(1)
%>
		  <tr>
			<th colspan=2>Transaction Response</th>
		  </tr>
		  <tr>
			<td align="right">Authorization&nbsp;#:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspAuthNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Success:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspSuccess").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Customer Tx #:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspCustTransNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Merchant Tx #:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspMerchTransNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">AVS Code:</td>
			<td align="left">&nbsp;<%= pstrAVSCode %></td>
		  </tr>
		  <tr>
			<td align="right">AUX Message:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspAUXMsg").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Action Code:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspActionCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Retrieval Code:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspRetrievalCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Error Message:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspErrorMsg").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Error Location:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspErrorLocation").Value %></td>
		  </tr>
<%
			objRSCC.MoveNext
		Next 'i
		Response.Write "</table>"
	End If	'objRSCC.RecordCount >= 1
		
End Sub	'ShowTransactionResponse_Complete

'**************************************************************************************************************************************************

Sub ShowTransactionResponse_short(byRef objRSCC)

Dim paryCodes
Dim pstrAVSCode
Dim pstrCCVCode

	If objRSCC.RecordCount >= 1 Then
		Response.Write "<table class=""tbl"" width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" rules=""none"">"
		For i = 1 To objRSCC.RecordCount
			paryCodes = Split(objRSCC.Fields("trnsrspAVSCode").Value & "", "|")
			If UBound(paryCodes) >= 0 Then pstrAVSCode = paryCodes(0)
			If UBound(paryCodes) >= 1 Then pstrCCVCode = paryCodes(1)
%>
		  <tr>
			<th colspan=2>Transaction Response</th>
		  </tr>
		  <tr>
			<td align="right">Authorization&nbsp;#:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspAuthNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Success:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspSuccess").Value %></td>
		  </tr>
		  <tr>
			<td align="right">AVS Code:</td>
			<td align="left">&nbsp;<%= pstrAVSCode %></td>
		  </tr>
		  <tr>
			<td align="right">CCV Code:</td>
			<td align="left">&nbsp;<%= pstrCCVCode %></td>
		  </tr>
		  <tr>
			<td align="right">Retrieval Code:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspRetrievalCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Error Message:</td>
			<td align="left">&nbsp;<%= objRSCC.Fields("trnsrspErrorMsg").Value %></td>
		  </tr>
<%
			objRSCC.MoveNext
		Next 'i
		Response.Write "</table>"
	End If	'objRSCC.RecordCount >= 1
	
End Sub	'ShowTransactionResponse_short

'**************************************************************************************************************************************************

Sub ShowCC(lngOrderID)

Dim pstrSQL
Dim pobjRSCC
Dim mstrCCNumber

	If cblnSQLDatabase Then
		pstrSQL = "SELECT sfCPayments.*, sfTransactionTypes.transName " _
				& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON convert(Integer,sfCPayments.payCardType) = sfTransactionTypes.transID" _
				& " Where sfCPayments.payCardType is not null AND OrderID=" & lngOrderID
	Else
		pstrSQL = "SELECT sfCPayments.*, sfTransactionTypes.transName " _
				& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON CLng(sfCPayments.payCardType) = sfTransactionTypes.transID" _
				& " Where  sfCPayments.payCardType is not null AND OrderID=" & lngOrderID
	End If
	Set pobjRSCC = GetRS(pstrSQL)

	If pobjRSCC.State = 0 Then
%>
		<table class="tbl" width="100%" cellpadding="0" cellspacing="0" border="0" rules="none">
		  <tr>
			<th>No transaction record is available for order <%= lngOrderID %></th>
		  </tr>
	    </table>
<%
	ElseIf pobjRSCC.EOF Then
%>
		<table class="tbl" width="100%" cellpadding="0" cellspacing="0" border="0" rules="none">
		  <tr>
			<th>No transaction record for <%= lngOrderID %></th>
		  </tr>
	    </table>
<%
	Else
		mstrCCNumber = DecryptCardNumber(pobjRSCC.Fields("payCardNumber").Value, False)
%>
		<table ID="tblCCInfo">
		  <tr>
			<th colspan=2><a href="" title="Click to Edit Card Information" onclick="OpenHelp('sfCPaymentsAdmin.asp?Action=viewItem&ViewID=<%= pobjRSCC.Fields("payID").Value %>'); return false;">Credit Card Information</a></td>
		  </tr>
		  <tr>
			<td align="right">Name on Credit Card:</td>
			<td align="left">&nbsp;<%= pobjRSCC.Fields("payCardName").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Credit Card Type:</td>
			<td align="left">&nbsp;<%= pobjRSCC.Fields("transName").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Credit Card Number:</td>
			<td align="left">&nbsp;<%= mstrCCNumber %></td>
		  </tr>
		  <%
			If Len(cstrCCV) > 0 Then
				On Error Resume Next
		  %>
		  <tr>
			<td align="right">CCV:</td>
			<td align="left">&nbsp;<%= pobjRSCC.Fields(cstrCCV).Value %></td>
		  </tr>
		  <%
				If Err.number <> 0 Then
					Response.Write "<h4><font color=red>You may have define a field for the Card Code (<em>" & cstrCCV & "</em>) in ssOrderAdmin_common.asp which is not in the database.</font></h4>"
					Err.Clear
				End If
			End If 'Len(cstrCCV) > 0
		  %>
		  <tr>
		    <td align="right">Expiration Date:</td>
		    <td align="left">&nbsp;<%= pobjRSCC.Fields("payCardExpires").Value %></td>
		  </tr>
	    </table>
<%
	End If	'pobjRSCC.EOF
	
	Call ReleaseObject(pobjRSCC)

End Sub 'ShowCC

'**************************************************************************************************************************************************

Function isAddressDifferent(byRef objRS)

Dim pstrBillingAddress
Dim pstrShippingAddress

	With objRS
		pstrBillingAddress = Trim(.Fields("custCompany").Value & "") _
						   & Trim(.Fields("custFirstName").Value & "") _
						   & Trim(.Fields("custMiddleInitial").Value & "") _
						   & Trim(.Fields("custLastName").Value & "") _
						   & Trim(.Fields("custAddr1").Value & "") _
						   & Trim(.Fields("custAddr2").Value & "") _
						   & Trim(.Fields("custCity").Value & "") _
						   & Trim(.Fields("custState").Value & "") _
						   & Trim(.Fields("custZip").Value & "") _
						   & Trim(.Fields("custCountry").Value & "")
	
		pstrBillingAddress = Trim(.Fields("cshpaddrShipCompany").Value & "") _
						   & Trim(.Fields("cshpaddrShipFirstName").Value & "") _
						   & Trim(.Fields("cshpaddrShipMiddleInitial").Value & "") _
						   & Trim(.Fields("cshpaddrShipLastName").Value & "") _
						   & Trim(.Fields("cshpaddrShipAddr1").Value & "") _
						   & Trim(.Fields("cshpaddrShipAddr2").Value & "") _
						   & Trim(.Fields("cshpaddrShipCity").Value & "") _
						   & Trim(.Fields("cshpaddrShipState").Value & "") _
						   & Trim(.Fields("cshpaddrShipZip").Value & "") _
						   & Trim(.Fields("cshpaddrShipCountry").Value & "")
	
	End With
	
	isAddressDifferent = CBool(LCase(pstrBillingAddress) <> LCase(pstrBillingAddress))

End Function	'isAddressDifferent

'**************************************************************************************************************************************************

Function getProductAttributeDetails(byVal strProductID, byVal strAttributeCategory, byVal strAttributeDetail, byVal aryFieldNames, byRef aryResult)

Dim pblnSuccess
Dim pobjCMD
Dim pobjRS
Dim pstrSQL
Dim i
Dim pstrFieldsToGet

	pblnSuccess = True
	If isArray(aryFieldNames) Then
		pstrFieldsToGet = aryFieldNames(0)
		For i = 1 To UBound(aryFieldNames)
			pstrFieldsToGet = pstrFieldsToGet & ", " & aryFieldNames(i)
		Next
	Else
		pstrFieldsToGet = aryFieldNames
	End If

	pstrSQL = "SELECT " & pstrFieldsToGet _
			& " FROM (sfProducts INNER JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) INNER JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
			& " WHERE ((sfAttributeDetail.attrdtName=?) AND (sfAttributes.attrName=?) AND (sfProducts.prodID=?))"

	Set pobjCMD = Server.CreateObject("ADODB.Command")
	With pobjCMD
		.ActiveConnection = cnn
		.CommandType = adCmdText
		.CommandText = pstrSQL

		'.Parameters.Append .CreateParameter("attrdtFileName", adVarChar, adParamInputOutput, 255, NULL)
		.Parameters.Append .CreateParameter("attrdtName", adVarChar, adParamInput, 255, strAttributeDetail)
		.Parameters.Append .CreateParameter("attrName", adVarChar, adParamInput, 255, strAttributeCategory)
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 255, strProductID)

		On Error Resume Next
		Set	pobjRS = .Execute

		If False Then
			Response.Write "<fieldset><legend></legend>"
			Response.Write "Product ID: " & strProductID & "<BR>"
			Response.Write "Attribute Category: " & strAttributeCategory & "<BR>"
			Response.Write "Attribute Detail: " & strAttributeDetail & "<BR>"
			Response.Write "pstrSQL: " & pstrSQL & "<BR>"
			Response.Write "pobjRS.EOF: " & pobjRS.EOF & "<BR>"
			Response.Write "</fieldset>"
		End If
		
		If Err.number <> 0 Then
			Response.Write "<fieldset><legend>Error in getProductAttributeDetails</legend>" _
						 & "Error " & Err.number & ": " & Err.Description & "<br />" & vbcrlf _
						 & "sql = " & pstrSQL & "</fieldset>" & vbcrlf
			Err.Clear
			pblnSuccess = False
		ElseIf Not pobjRS.EOF Then
			If isArray(aryFieldNames) Then
				For i = 0 To UBound(aryFieldNames)
					aryResult(i) = Trim(pobjRS.Fields(aryFieldNames(i)).Value & "")
				Next
			Else
				aryResult = Trim(pobjRS.Fields(aryFieldNames).Value & "")
			End If
		End If
		Call ReleaseObject(pobjRS)
		
	End With	'pobjCMD
	
	Call ReleaseObject(pobjCMD)
	
	getProductAttributeDetails = pblnSuccess

End Function	'getProductAttributeDetails
%>