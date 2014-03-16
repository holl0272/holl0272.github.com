<%
'********************************************************************************
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************
'
'	This page is separated from ssclsCartContents.asp to enable easier editing
'   It is located inside the class module, Public Sub displayOrder_CheckoutView
'
'	The html fragment on this page MUST be a self-contained table
'
'**********************************************************
'*	Sub routine variables
'**********************************************************
	
	Dim i
	Dim j
	Dim paryOrderItem
	Dim paryAttribute
	Dim pstrFontClass
	Dim pstrAttributeName
	Dim pstrImageSRC
	Dim pstrProdLink
	Dim plngProdAttrNum

	If pblnEmptyCart Then
	%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="5" width="40%" class="tdAltFont1"><center><p style="margin-top:25pt"><font class='TopBanner_Large'><b>No items in shopping cart</b></font></p><br />Please press <a href="<%= getLastSearch %>"><img src="<%= C_BTN04 %>" border="0" name="continue_search" alt="Continue Search" ></a> to begin searching for items.</center></td>			
	</tr>
	</table>				
	<%
	Else
	%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="1%" align="left" class="tdContentBar"></td>
		<td width="54%" align="left" class="tdContentBar">Item</td>
		<td width="15%" align="center" class="tdContentBar">Qty</td>
		<td width="15%" align="center" class="tdContentBar">Unit Price</td>
		<td width="15%" align="center" class="tdContentBar">Price</td>
	</tr>
	<%
		For i = 0 To plngUniqueOrderItemCount
			paryOrderItem = paryOrderItems(i)

			If (i mod 2) = 1 Then 
				pstrFontClass="tdAltFont1"
			Else 	
				pstrFontClass="tdAltFont2"
			End If
			
			If Len(CStr(cbytOrderViewImageSrc)) > 0 Then
				pstrImageSRC = paryOrderItem(cbytOrderViewImageSrc)
			Else
				pstrImageSRC = ""
			End If
					
			If Len(pstrImageSRC) > 0 And LCase(Left(pstrImageSRC, 6)) <> "https:" Then pstrImageSRC = "../" & pstrImageSRC
			
			pstrProdLink = paryOrderItem(enOrderItem_prodLink)
			If Len(pstrProdLink) = 0 Then pstrProdLink = "detail.asp?product_id=" & paryOrderItem(enOrderItem_prodID)
			If LCase(Left(pstrProdLink, 5)) <> "http:" Then pstrProdLink = adminDomainName & pstrProdLink
	%>
	<tr>
		<td class="<%= pstrFontClass %>" valign="top" align="left"><% If Len(pstrImageSRC) > 0 Then %><a href="<%= pstrProdLink %>"><img src="<%= pstrImageSRC %>" alt="<%= stripHTML(paryOrderItem(enOrderItem_prodName)) %>" /></a><% End If 'Len(pstrImageSRC) > 0 %></td>
		<td valign="top" class="<%= pstrFontClass %>" align="left">
			<a href="<%= pstrProdLink %>"><b><%= paryOrderItem(enOrderItem_prodName) %></b></a><br />
			<%
			If Len(cstrCartDisplay_ProductID) > 0 Then Response.Write cstrCartDisplay_ProductID & paryOrderItem(enOrderItem_prodID) & "<br />"
			If Len(cstrCartDisplay_MfgName) > 0 Then Response.Write cstrCartDisplay_MfgName & ManufacturerName(paryOrderItem) & "<br />"

			plngProdAttrNum = paryOrderItem(enOrderItem_AttributeCount)
			If paryOrderItem(enOrderItem_AttributeCount) > 0 Then
				paryAttributes = paryOrderItem(enOrderItem_AttributeArray)
				For j = 0 To paryOrderItem(enOrderItem_AttributeCount) - 1
					pstrAttributeName = Trim(paryAttributes(j)(enAttributeItem_attrName))
					If Right(pstrAttributeName, 1) = ":" Then pstrAttributeName = Left(pstrAttributeName, Len(pstrAttributeName) - 1)
			%>
			&nbsp;&nbsp;<%= pstrAttributeName %>: <%= paryAttributes(j)(enAttributeItem_attrdtName) %><br />
			<%
				Next 'j
			End If	'isArray(paryOrderItemAttributes)
			%>
		</td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><%= paryOrderItem(enOrderItem_odrdttmpQuantity) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_UnitPrice)) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_odrdttmpQuantity) * paryOrderItem(enOrderItem_UnitPrice)) %></td>
	</tr> 

	<% If paryOrderItem(enOrderItem_prodSetupFeeOneTime) > 0 Then %>
	<tr>
		<td align="left" class="<%= pstrFontClass %>" valign="top">&nbsp;</td>	  
		<td class="<%= pstrFontClass %>" align="left"><b>Set-up Fee (Per product)</b></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top">1</td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_prodSetupFeeOneTime)) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_prodSetupFeeOneTime)) %></td>
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_gwActivate) %>

	<% If paryOrderItem(enOrderItem_prodSetupFee) > 0 Then %>
	<tr>
		<td align="left" class="<%= pstrFontClass %>" valign="top">&nbsp;</td>	  
		<td class="<%= pstrFontClass %>" align="left"><b>Set-up Fee (Each product)</b></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top">1</td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_prodSetupFee)) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_prodSetupFee)) %></td>
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_gwActivate) %>

	<% If paryOrderItem(enOrderItem_gwActivate) = 1 Then %>
	<tr>
		<td align="center" class="<%= pstrFontClass %>" valign="top">&nbsp;</td>	  
		<td class="<%= pstrFontClass %>" align="left"><b>Gift Wrap</b></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><%= paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_gwPrice)) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) * paryOrderItem(enOrderItem_gwPrice)) %></td>
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_gwActivate) %>

	<% If paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 Then %>
	<tr>
		<td class="<%= pstrFontClass %>" valign="top">&nbsp;</td>	  
		<td colspan="4" align="left" class="<%= pstrFontClass %>" valign="top"><b>Back Ordered Qty: <%= paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) %></b> (of the above items)<br />Backordered items will be included in the order total but will not be billed until shipped.</td>	  
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 %>
	
	<% If i <> plngUniqueOrderItemCount Then	'Do not add spacer row after last item %>
	<tr><td colspan="5" height="12px">&nbsp;</td></tr>
	<% End If	'plngUniqueOrderItemCount %>	
	<%
		Next 'i
	%>

	<tr>
	  <td colspan="5" valign="middle"><hr /></td>
	</tr>

	<tr>
	  <td class="tdContent2" colspan="2" valign="top" align="left" width="60%">
		<% 'Call displayVisitorShippingPreferences %>
	  </td>
	  <td class="tdContent2" colspan="3" valign="top" align="right" width="40%">
		<% Call ShowOrderDiscounts %>
		<% Call displayOrderSummary %>
	  </td>
	</tr>
	</table>
	<%
	End If	'pblnEmptyCart
	%>