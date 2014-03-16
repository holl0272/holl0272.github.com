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
'   It is located inside the class module, Public Sub displayOrder
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
	Dim pstrAttributeName
	Dim pstrFontClass
	Dim pstrImageSRC
	Dim pstrProdLink
	Dim plngProdAttrNum
	Dim pstrMinimumOrderMessage_Temp

	If pblnEmptyCart Then
	%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="5" width="40%" class="tdAltFont1" align="center">
		  <p style="margin-top:25pt"><font class='TopBanner_Large'><b>No items in shopping cart</b></font></p>
		  <%
		  If IsSaveCartActive = 1 Then
			If hasSavedCart(visitorLoggedInCustomerID) Then
			%>
			<p><a href="savecart.asp" title="View Saved Cart"><img class="inputImage" src="<%= C_BTN08 %>" alt="View Wish List"></a></p>
			<%  
			End If
		  End If
		  %>
		  Please press <a href="<%= getLastSearch %>"><img class="inputImage" src="<%= C_BTN04 %>" name="continue_search" alt="Continue Search"></a> to begin searching for items.
		</td>			
	</tr>
	</table>				
	<%
	Else
	%>
	<form method="POST" name="frmQty" id="frmQty" action="order.asp" onSubmit="if(this.recalc.value != 'no'){this.recalc.value=1;}">
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
			
			pstrProdLink = paryOrderItem(enOrderItem_prodLink)
			If Len(pstrProdLink) = 0 Then pstrProdLink = "detail.asp?product_id=" & paryOrderItem(enOrderItem_prodID)
	%>
	<tr>
		<td class="<%= pstrFontClass %>" valign="top" align="left"><% If Len(pstrImageSRC) > 0 Then %><a href="<%= pstrProdLink %>"><img src="<%= pstrImageSRC %>" alt="<%= stripHTML(paryOrderItem(enOrderItem_prodName)) %>" /></a><% End If 'Len(pstrImageSRC) > 0 %></td>
		<td valign="top" class="<%= pstrFontClass %>" align="left">
			<a href="<%= pstrProdLink %>"><b><%= paryOrderItem(enOrderItem_prodName) %></b></a><br />
			<%
			If Len(cstrCartDisplay_ProductID) > 0 Then Response.Write cstrCartDisplay_ProductID & " " & paryOrderItem(enOrderItem_prodID) & "<br />"
			If Len(cstrCartDisplay_MfgName) > 0 Then Response.Write cstrCartDisplay_MfgName & " " & ManufacturerName(paryOrderItem) & "<br />"

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
			<input type="hidden" name="sProdID<%= i %>" ID="sProdID<%= i %>" value="<%= paryOrderItem(enOrderItem_prodID) %>">
			<input type="hidden" name="iOrderID<%= i %>" ID="iOrderID<%= i %>" value="<%= paryOrderItem(enOrderItem_tmpID) %>">
			<input type="hidden" name="iQuantity<%= i %>" ID="iQuantity<%= i %>" value="<%= paryOrderItem(enOrderItem_odrdttmpQuantity) %>">
			<input type="hidden" name="iProdAttrNum<%= i %>" ID="iProdAttrNum<%= i %>" value="<%= plngProdAttrNum %>">
			<p>
			<input type="image" class="inputImage" src="<%= C_BTN06 %>" onmousedown="recalc.value='no';" name="DeleteFromOrder<%= i %>" ID="DeleteFromOrder<%= i %>">
			<% If IsSaveCartActive = 1 then %>
			&nbsp;<input type="image" class="inputImage" src="<%= C_BTN07 %>" name="SaveToCart<%= i%>" ID="SaveToCart<%= i%>">
			<% end if %>
			</p>
		</td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><input type="text" class="formDesign" size="2" name="FormQuantity<%= i %>" ID="FormQuantity<%= i %>" value="<%= paryOrderItem(enOrderItem_odrdttmpQuantity) %>"></td>
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
		<td align="left" class="<%= pstrFontClass %>" valign="top">&nbsp;</td>	  
		<td class="<%= pstrFontClass %>" align="left"><b>Gift Wrap</b></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><input type="text" class="formDesign" size="2" name="GWQTY<%= i %>" ID="GWQTY<%= i %>" value="<%= paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) %>"></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_gwPrice)) %></td>
		<td align="center" class="<%= pstrFontClass %>" valign="top"><% writeCustomCurrency(paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) * paryOrderItem(enOrderItem_gwPrice)) %></td>
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_gwActivate) %>

	<% If paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 Then %>
	<tr>
		<td class="<%= pstrFontClass %>" valign="top" align="left">&nbsp;</td>	  
		<td colspan="4" align="left" class="<%= pstrFontClass %>" valign="top"><b>Back Ordered Qty: <%= paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) %></b> (of the above items)<br />Backordered items will be included in the order total but will not be billed until shipped.</td>	  
	</tr> 
	<% End If	'paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 %>
	
	<% If i <> plngUniqueOrderItemCount Then	'Do not add spacer row after last item %>
	<tr><td colspan="5" height="12px">&nbsp;</td></tr>
	<% End If	'plngUniqueOrderItemCount %>	
	<%
		Next 'i
	%>
	<% If adminFreeShippingIsActive And adminShipType = 1 Then %>
	<tr>
		<td width="40%"></td>
		<td width="15%" align="center" valign="top"></td>
		<td nowrap colspan="2" width="30%" align="right" valign="top"><font class="Error">Free Shipping on orders over <b><% writeCustomCurrency(adminFreeShippingAmount) %></b>!</font></td>
		<td width="15%" align="center" valign="top"></td>
	</tr> 
	<% End If	'adminFreeShippingIsActive %>
	<tr>
	  <td colspan="5" valign="middle"><hr /></td>
	</tr>
	<tr>
		<td colspan="3" align="right" valign="top">
	      <font class="Error">&nbsp;</font>
	    </td>
		<td colspan="2" align="right" valign="top">
	      <input type="hidden" name="iProductCounter" ID="iProductCounter" value="<%= i - 1 %>">
	      <input type="hidden" name="recalc" ID="recalc" value="" />
	      <input type="image" class="inputImage" src="<%= C_BTN14 %>" name="Recalculate" id="Recalculate" value="" onclick="javascript:recalc.value='';" /> 
	      <input type="image" class="inputImage" src="images/buttons/saveAllItems.gif" name="MoveAll" id="MoveAll"> 
	      <% If pdblSubTotalWithDiscount >= pdblMinimumOrderAmount Then %>
	      <br /><input type="image" class="inputImage" src="<%= C_BTN05 %>" name="checkout" onclick="document.frmCheckout.submit(); return false;"><br />
	      <% End If %>
		</td>
	</tr>
	</table>
	</form>
	<div align="center"><hr /></div>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
	  <td class="tdContent2" valign="top" align="left">
		<% Call displayVisitorShippingPreferences %>
	  </td>
	  <td class="tdContent2" valign="top" align="right">
		<script language="javascript" type="text/javascript">
		
			var blnPromotionSubmitted = false;
			function SubmitPromotion()
			{
				
				var theForm = document.frmPromo;
				if (blnPromotionSubmitted)
				{
					theForm.btnSubmitPromo.value = "Please be patient . . .";
					return false;
				}
				
				if (theForm.PromoCode.value == "")
				{
					alert("Please enter a promotion code.")
					theForm.PromoCode.focus();
					return false;
				}
				
				theForm.btnSubmitPromo.value = "Retrieving . . .";
				
				blnPromotionSubmitted = true;

				return true;
				
			}
			
			var blnCertificateSubmitted = false;
			function SubmitCertificate()
			{
				var theForm = document.frmGCRegister;
				if (blnCertificateSubmitted)
				{
					theForm.btnSubmitGC.value = "Please be patient . . .";
					return false;
				}
				
				if (theForm.Certificate.value=='')
				{
					alert('Please enter a certificate number.');
					document.frmGCRegister.Certificate.focus();
					return false;
				}
				
				theForm.btnSubmitGC.value = "Retrieving . . .";
				
				blnCertificateSubmitted = true;

				return true;
			}
		</script>
		<div id="divMainCheckout">
		<% Call ShowOrderDiscounts %>
		<% Call displayOrderSummary %>
		<div id="divCheckoutWrapper">
            <%If Not pblnEmptyCart Then %>                      
              <form action="<%= C_SecurePath %>" method="post" name="frmCheckout" ID="frmCheckout">
                <input type="hidden" name="SessionID" ID="SessionID" value="<%= SessionID %>">
                <% If pdblSubTotalWithDiscount >= pdblMinimumOrderAmount Then %>
                <input type="image" class="inputImage" src="<%= C_BTN05 %>" name="checkout"><br />
                <font class="Content_Small">You will go to our secure server to finish your transaction.</font><br />
                <!--<input type=image NAME="btn_xpressCheckout" ID="btn_xpressCheckout" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif"><br />-->
                <% Else %>
                <%
					pstrMinimumOrderMessage_Temp  = pstrMinimumOrderMessage
					pstrMinimumOrderMessage_Temp = Replace(pstrMinimumOrderMessage_Temp, "{MinimumOrderAmount}", customCurrency(pdblMinimumOrderAmount))
					pstrMinimumOrderMessage_Temp = Replace(pstrMinimumOrderMessage_Temp, "{amountShort}", customCurrency(pdblMinimumOrderAmount - pdblSubTotalWithDiscount))
					Response.Write pstrMinimumOrderMessage_Temp
                %>
                <% End If %>
              </form>
            <%End If	'Not pblnEmptyCart %>
			<form action="order.asp" id="frmPromo" name="frmPromo" method="post" onsubmit="return(SubmitPromotion());">
			  <table border="0" cellspacing="0" cellpadding="2" class="tdContent2">
				<% If len(mstrPromotionRegistrationMessage) > 0 Then %>
				<tr>
				  <td align="center"><%= mstrPromotionRegistrationMessage %></td>
				</tr>
				<% End If %>
				<tr>
				  <td align="center" nowrap><input id="PromoCode" name="PromoCode" value="<%= mstrPromotionCode %>">&nbsp;<input id=btnSubmitPromo name=btnSubmitPromo type=submit value=Enter><br />
				  Do you have a promotion code?</td>
				</tr>
			</table>
			</form>
<!--
			<form ID="frmGCRegister" Name="frmGCRegister" action="order.asp" method="post" onsubmit="return(SubmitCertificate());">
			<table border="0" cellspacing="0" cellpadding="2" class="tdContent2">
				<% If len(mstrGiftCertificateRegistrationMessage) > 0 Then %>
				<tr>
				  <td align="center"><font class="Error"><%= mstrGiftCertificateRegistrationMessage %></font></td>
				</tr>
				<% End If %>
				<tr>
				  <td align="center" nowrap><input id="Certificate" name="Certificate" value="<%= mstrCertificate %>">&nbsp;<input id=btnSubmitGC name=btnSubmitGC type=submit value=Enter><br />
				  Do you have a gift card to redeem?</td>
				</tr>
			</table>
			</form>
-->
              <%
				If IsSaveCartActive = 1 Then
					If hasSavedCart(visitorLoggedInCustomerID) Then
					%>
					<p><a href="savecart.asp" title="View Saved Cart"><img class="inputImage" src="<%= C_BTN08 %>" alt="View Wish List"></a></p>
					<%  
					End If
				End If
              %>
              <a href="<%= getLastSearch %>"><img class="inputImage" src="<%= C_BTN04 %>" name="continue_search" alt="Continue Search" ></a>
		</div>
		</div>
	  </td>
	</tr>
	</table>
	<%
	End If	'pblnEmptyCart
	%>