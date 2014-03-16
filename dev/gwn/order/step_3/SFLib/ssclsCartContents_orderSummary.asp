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
'   It is located inside the class module, Public Sub displayOrderSummary
'
'	The html fragment on this page MUST be a self-contained table
'
'**********************************************************
'*	Sub routine variables
'**********************************************************

' NONE
	
%>
<div id="divOrderSummaryWrapper">
  <table border="0" cellspacing="0" cellpadding="2" class="tdContent2" width="100%">
	<tr>
		<th colspan="2" class="tdTopBanner" nowrap>Order Summary:</th>
	</tr>
	<% If pdblDiscount > 0 Then %>
	<tr><td width="75%" align="right">Product Sub Total: </td><td align="right"><% writeCustomCurrency(pdblSubTotal) %></td></tr>
	<tr><td width="75%" align="right">Discount: </td><td align="right"><% writeCustomCurrency(pdblDiscount) %></td></tr>
	<% End If	'pdblDiscount > 0 %>
	<tr><td width="75%" align="right">Sub Total: </td><td align="right"><% writeCustomCurrency(pdblSubTotalWithDiscount) %></td></tr>
	<% If pdblHandling > 0 Then %>
	<tr><td width="75%" align="right">Handling: </td><td align="right"><% writeCustomCurrency(pdblHandling) %></td></tr>
	<% End If	'pdblHandling > 0 %>
	<% If pdblCOD > 0 Then %>
	<tr><td width="75%" align="right">COD: </td><td align="right"><% writeCustomCurrency(pdblCOD) %></td></tr>
	<% End If	'pdblCOD > 0 %>
	<% If pblnOrderIsShipped Then %>
	<tr><td width="75%" align="right" nowrap>Shipping (<%= pstrShipMethodName %>): </td><td align="right"><% writeCustomCurrency(pdblShipping) %></td></tr>
	<% End If	'pblnOrderIsShipped %>
	<% If pdblLocalTax > 0 Then %>
	<tr><td width="75%" align="right">Local Tax: </td><td align="right"><% writeCustomCurrency(pdblLocalTax) %></td></tr>
	<% End If	'pdblLocalTax > 0 %>
	<% If pdblStateTax > 0 Then %>
	<tr><td width="75%" align="right">State Tax: </td><td align="right"><% writeCustomCurrency(pdblStateTax) %></td></tr>
	<% End If	'pdblStateTax > 0 %>
	<% If pdblCountryTax > 0 Then %>
	<tr><td width="75%" align="right">Country Tax: </td><td align="right"><% writeCustomCurrency(pdblCountryTax) %></td></tr>
	<% End If	'pdblCountryTax > 0 %>
	<tr><td width="75%" align="right">Cart Total: </td><td align="right"><% writeCustomCurrency(pdblCartTotal) %></td></tr>
	<% If pdblAvailableStoreCredit > 0 Then %>
	<tr><td width="75%" align="right">Credit: </td><td align="right"><% writeCustomCurrency(pdblAvailableStoreCredit) %></td></tr>
	<tr><td width="75%" align="right">Amount Due: </td><td align="right"><% writeCustomCurrency(pdblAmountDue) %></td></tr>
	<% End If	'pdblStoreCredit > 0 %>
  </table>
</div>
