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
'	The html fragment on this page MUST be a self-contained element such as a table, p, div, span, etc.
'
'**********************************************************
'*	Sub routine variables
'**********************************************************

' NONE
	
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" id="miniCartOrderSummary">
	<tr><th colspan="2" class="cartSummaryBanner">Order Summary:</th></tr>
	<% If pdblDiscount > 0 Then %>
	<tr><td class="cartSummaryLabel">Product Sub Total: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblSubTotal) %></td></tr>
	<tr><td class="cartSummaryLabel">Discount: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblDiscount) %></td></tr>
	<% End If	'pdblDiscount > 0 %>
	<tr><td  class="cartSummaryLabel">Sub Total: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblSubTotalWithDiscount) %></td></tr>
	<% If pdblHandling > 0 Then %>
	<tr><td class="cartSummaryLabel">Handling: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblHandling) %></td></tr>
	<% End If	'pdblHandling > 0 %>
	<% If pdblCOD > 0 Then %>
	<tr><td class="cartSummaryLabel">COD: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblCOD) %></td></tr>
	<% End If	'pdblCOD > 0 %>
	<% If pblnOrderIsShipped Then %>
	<tr><td class="cartSummaryLabel" nowrap>Shipping: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblShipping) %></td></tr>
	<% End If	'pblnOrderIsShipped %>
	<% If pdblLocalTax > 0 Then %>
	<tr><td class="cartSummaryLabel">Local Tax: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblLocalTax) %></td></tr>
	<% End If	'pdblLocalTax > 0 %>
	<% If pdblStateTax > 0 Then %>
	<tr><td class="cartSummaryLabel">State Tax: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblStateTax) %></td></tr>
	<% End If	'pdblStateTax > 0 %>
	<% If pdblCountryTax > 0 Then %>
	<tr><td class="cartSummaryLabel">Country Tax: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblCountryTax) %></td></tr>
	<% End If	'pdblCountryTax > 0 %>
	<tr><td class="cartSummaryLabel">Cart Total: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblCartTotal) %></td></tr>
	<% If pdblAvailableStoreCredit > 0 Then %>
	<tr><td class="cartSummaryLabel">Credit: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblAvailableStoreCredit) %></td></tr>
	<tr><td class="cartSummaryLabel">Amount Due: </td><td class="cartSummaryContent"><% writeCustomCurrency(pdblAmountDue) %></td></tr>
	<% End If	'pdblStoreCredit > 0 %>
</table>
