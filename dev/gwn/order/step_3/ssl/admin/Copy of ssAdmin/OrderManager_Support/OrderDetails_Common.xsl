<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<!--
This file contains supporting Common variables and templates

Common variables:

	
Common templates:

-->
<xsl:template name="paymentInformation">

  <tr class="tblhdr">
    <th colspan="2">Payment Information</th>
  </tr>
  <tr>
    <td valign="top">
      <table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="1" cellpadding="0" cellspacing="0">
		<colgroup align="right" />
		<colgroup align="left" />
		
		<xsl:choose>
			<xsl:when test="PayMethod=1"><!-- eCheck -->
				<xsl:if test="string-length(Payments/CheckNumber)>0">
				<tr><td>Payment method:</td><td>&amp;nbsp;eCheck</td></tr>
				<tr><td>Bank Name:</td><td>&amp;nbsp;<xsl:value-of select="Payments/BankName" /></td></tr>
				<tr><td>Routing Number:</td><td>&amp;nbsp;<xsl:value-of select="Payments/RoutingNumber" /></td></tr>
				<tr><td>Account Number:</td><td>&amp;nbsp;<xsl:value-of select="Payments/AccountNumber" /></td></tr>
				<tr><td>Check #:</td><td>&amp;nbsp;<xsl:value-of select="Payments/CheckNumber" /></td></tr>
				</xsl:if>
			</xsl:when>
			<xsl:when test="PayMethod=2"><!-- COD -->
				<tr><td>Payment method:</td><td>&amp;nbsp;COD</td></tr>
			</xsl:when>
			<xsl:when test="PayMethod=3"><!-- PO -->
				<xsl:if test="string-length(Payments/PONumber)>0">
				<tr><td>Payment method:</td><td>&amp;nbsp;Purchase Order</td></tr>
				<tr><td>PO #:</td><td>&amp;nbsp;<xsl:value-of select="Payments/PONumber" /></td></tr>
				</xsl:if>
			</xsl:when>
			<xsl:when test="PayMethod=4 or PayMethod=5"><!-- PhoneFax: Recorded, PhoneFax: Non-Recorded -->
				<tr><td>Payment method:</td><td>&amp;nbsp;<xsl:value-of select="PayMethodName" /></td></tr>
				<xsl:if test="string-length(Payments/CardType)>0">
					<tr><td>Card Name:</td><td>&amp;nbsp;<xsl:call-template name="outputName"><xsl:with-param name='currentNode' select='billingAddress' /></xsl:call-template></td></tr>
					<tr><td>Card Type:</td><td>&amp;nbsp;<xsl:value-of select="Payments/CardType" /></td></tr>
					<tr><td>Card #:</td><td>&amp;nbsp;
						<xsl:choose>
							<xsl:when test="string-length(Payments/CreditCardNumber)>0"><xsl:value-of select="Payments/CreditCardNumber" /></xsl:when>
							<xsl:otherwise>****-****-****-****-<xsl:value-of select="Payments/Last4Digits" /></xsl:otherwise>
						</xsl:choose>
					</td></tr>
					<tr><td>Card Exp.:</td><td>&amp;nbsp;<xsl:value-of select="Payments/ExpireMonth" /> / <xsl:value-of select="Payments/ExpireYear" /></td></tr>
				</xsl:if>
			</xsl:when>
			<xsl:when test="PayMethod=6"><!-- PayPal: PayPal -->
				<tr><td>Payment method:</td><td>&amp;nbsp;PayPal</td></tr>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="string-length(Payments/CardType)>0">
				<tr><td>Payment method:</td><td>&amp;nbsp;Credit Card</td></tr>
				<tr><td>Card Name:</td><td>&amp;nbsp;<xsl:call-template name="outputName"><xsl:with-param name='currentNode' select='billingAddress' /></xsl:call-template></td></tr>
				<tr><td>Card Type:</td><td>&amp;nbsp;<xsl:value-of select="Payments/CardType" /></td></tr>
				<tr><td>Card #:</td><td>&amp;nbsp;
					<xsl:choose>
						<xsl:when test="string-length(Payments/CreditCardNumber)>0"><xsl:value-of select="Payments/CreditCardNumber" /></xsl:when>
						<xsl:otherwise>****-****-****-****-<xsl:value-of select="Payments/Last4Digits" /></xsl:otherwise>
					</xsl:choose>
				</td></tr>
				<tr><td>Card Exp.:</td><td>&amp;nbsp;<xsl:value-of select="Payments/ExpireMonth" /> / <xsl:value-of select="Payments/ExpireYear" /></td></tr>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>

      </table>
    </td>
    <td valign="top">
	  <xsl:for-each select="ProcessorResponse">
      <table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="1" cellpadding="0" cellspacing="0">
		<colgroup align="right" />
		<colgroup align="left" />
		<tr>
		  <th colspan="2" align="center">Transaction Response</th>
		</tr>
		<tr>
		  <td>Order&amp;nbsp;uid:</td>
		  <td>&amp;nbsp;<xsl:value-of select="OrderID" /></td>
		</tr>
		<tr>
		  <td>Authorization&amp;nbsp;#:</td>
		  <td>&amp;nbsp;<xsl:value-of select="AuthorizationNo" /></td>
		</tr>
		<tr>
		  <td>Success:</td>
		  <td>&amp;nbsp;<xsl:value-of select="Success" /></td>
		</tr>
		<tr>
		  <td>Customer Tx #:</td>
		  <td>&amp;nbsp;<xsl:value-of select="CustTransNo" /></td>
		</tr>
		<tr>
		  <td>Merchant Tx #:</td>
		  <td>&amp;nbsp;<xsl:value-of select="MerchantTransNo" /></td>
		</tr>
		<tr>
		  <td>AVS Code:</td>
		  <td>&amp;nbsp;<xsl:value-of select="AVSResult" /></td>
		</tr>
		<tr>
		  <td>CVV Result:</td>
		  <td>&amp;nbsp;<xsl:value-of select="CVVResult" /></td>
		</tr>
		<tr>
		  <td>Action Code:</td>
		  <td>&amp;nbsp;<xsl:value-of select="ActionCode" /></td>
		</tr>
		  <tr>
		  <td>Retrieval Code:</td>
		  <td>&amp;nbsp;<xsl:value-of select="RetrievalCode" /></td>
		</tr>
		<tr>
		  <td>Error Message:</td>
		  <td>&amp;nbsp;<xsl:value-of select="ErrorMessage" /></td>
		</tr>
		<tr>
		  <td>Aux Message:</td>
		  <td>&amp;nbsp;<xsl:value-of select="AuxMessage" /></td>
		</tr>
		<tr>
		  <td>Error Location:</td>
		  <td>&amp;nbsp;<xsl:value-of select="ErrorLocation" /></td>
		</tr>
      </table>
	  </xsl:for-each>
    </td>
  </tr>
  </xsl:template>

<xsl:template name="orderStatus">
	<xsl:param name='shippingAddressCount' />

  <tr class="tblhdr">
    <th colspan="2">Order Status</th>
  </tr>
  <tr>
    <td valign="top" colspan="2">
      <table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="1" cellpadding="0" cellspacing="0" width="100%">
		<colgroup align="right" />
		<colgroup align="left" />
      <input type="hidden" name="PaymentsPending" id="PaymentsPending"><xsl:attribute name="value"><xsl:value-of select="PaymentsPending" /></xsl:attribute></input>
      <input type="hidden" name="BOPaymentsPending" id="BOPaymentsPending"><xsl:attribute name="value"><xsl:value-of select="BOPaymentsPending" /></xsl:attribute></input>
		  <tr><th>&amp;nbsp;</th><th>&amp;nbsp;Order <xsl:value-of select="OrderNumber" /> placed on <xsl:value-of select="DateOrdered" /></th></tr>
		  <xsl:if test="REMOTE_ADDR != ''">
		  <tr>
			<td>Remote IP:&amp;nbsp;</td>
			<td><xsl:value-of select="REMOTE_ADDR" />
				&amp;nbsp;<a target="whois">
					<xsl:attribute name="href">http://www.whois.sc/<xsl:value-of select="REMOTE_ADDR" /></xsl:attribute>
					<xsl:text>(Whois Lookup)</xsl:text>
				</a>
				&amp;nbsp;<a target="ReverseDNS">
					<xsl:attribute name="href">http://www.dnsstuff.com/tools/ptr.ch?ip=<xsl:value-of select="REMOTE_ADDR" /></xsl:attribute>
					<xsl:text>(Reverse DNS Lookup)</xsl:text>
				</a>
				&amp;nbsp;<a target="NetGeo">
					<xsl:attribute name="href">http://www.dnsstuff.com/tools/netgeo.ch?ip=<xsl:value-of select="REMOTE_ADDR" /></xsl:attribute>
					<xsl:text>(NetGeo Lookup)</xsl:text>
				</a>
			</td>
		  </tr>
		  </xsl:if>
		  <tr>
		    <td valign="top">Order Status:</td>
		    <td>&amp;nbsp;<input type="checkbox" name="Void" id="Void" value="1">
				<xsl:attribute name="onclick">setOrderNonPriceDirty();</xsl:attribute>
				<xsl:if test="Void=1">
				<xsl:attribute name="checked">checked</xsl:attribute>
				</xsl:if>
				</input>&amp;nbsp;<label for="Void">
				<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>Order has been voided
				</label><br />
		    
		    &amp;nbsp;<input type="checkbox" name="ssOrderFlagged" id="ssOrderFlagged" value="1">
				<xsl:attribute name="onclick">setOrderNonPriceDirty();</xsl:attribute>
				<xsl:if test="ssOrderFlagged=1">
				<xsl:attribute name="checked">checked</xsl:attribute>
				</xsl:if>
				</input>&amp;nbsp;<label for="ssOrderFlagged">
				<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>Flag this order
				</label>
		    </td>
		  </tr>

		  <tr>
		    <td>Payment Status:</td>
		    <td>&amp;nbsp;<input type="checkbox" name="dummyPaymentsPending" id="dummyPaymentsPending">
			<xsl:attribute name="onclick">var myValue; if(this.checked){myValue=0;}else{myValue=1;}frmData.PaymentsPending.value=myValue;setOrderNonPriceDirty();</xsl:attribute>
			<xsl:if test="PaymentsPending=0">
			<xsl:attribute name="checked">checked</xsl:attribute>
			</xsl:if>
			</input>&amp;nbsp;<label for="dummyPaymentsPending">
				<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>Payment has been collected
				</label>
		    
		    <br />&amp;nbsp;<input type="checkbox" name="dummyBOPaymentsPending" id="dummyBOPaymentsPending">
			<xsl:attribute name="onclick">var myValue; if(this.checked){myValue=0;}else{myValue=1;}frmData.BOPaymentsPending.value=myValue;setOrderNonPriceDirty();</xsl:attribute>
			<xsl:if test="BOPaymentsPending=0">
			<xsl:attribute name="checked">checked</xsl:attribute>
			</xsl:if>
			</input>&amp;nbsp;<label for="dummyBOPaymentsPending">
				<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>Backorder payment has been collected
				</label>
		    </td>
		  </tr>
		  
		  <tr>
		    <td valign="top">Processing Status:</td>
		    <td>
				<fieldset onmouseover="showHideElement(document.all('divProcessingStatuses'));" onmouseout="showHideElement(document.all('divProcessingStatuses'));">
					<xsl:variable name="ssOrderStatus"><xsl:value-of select="ssOrderStatus" /></xsl:variable>
					<legend>
						<xsl:for-each select="../OrderStatusOptions/OrderStatusOption">
								<xsl:if test="(@value = $ssOrderStatus) or (string-length($ssOrderStatus)=0 and position()=1)">
									<xsl:if test="string-length(@imgSRC)>0"><img border="0"><xsl:attribute name="src"><xsl:value-of select="@imgSRC" /></xsl:attribute></img></xsl:if>
									<span>
										<xsl:attribute name="class"><xsl:value-of select="@class" /></xsl:attribute>
										<xsl:value-of select="@text" />
									</span>
								</xsl:if>
						</xsl:for-each>
					</legend>
					<div id="divProcessingStatuses" style="display:none;">
						<xsl:for-each select="../OrderStatusOptions/OrderStatusOption">
							<xsl:if test="string-length(@imgSRC)>0"><img border="0"><xsl:attribute name="src"><xsl:value-of select="@imgSRC" /></xsl:attribute></img></xsl:if>

							<input type="radio" name="ssOrderStatus" onclick="setOrderNonPriceDirty();">
								<xsl:attribute name="id">ssOrderStatus<xsl:value-of select="position()-1" /></xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="@value" /></xsl:attribute>
								<xsl:if test="(@value = $ssOrderStatus) or (string-length($ssOrderStatus)=0 and position()=1)"><xsl:attribute name="checked" /></xsl:if>
							</input>
						
							<xsl:text>&amp;nbsp;</xsl:text>
						
							<span>
								<xsl:attribute name="class"><xsl:value-of select="@class" /></xsl:attribute>
								<label>
									<xsl:attribute name="for">ssOrderStatus<xsl:value-of select="position()-1" /></xsl:attribute>
									<xsl:value-of select="@text" />
								</label><br />
							</span>
						</xsl:for-each>
					</div>
				</fieldset>
		    &amp;nbsp;<input type="checkbox" name="ssExported" id="ssExported" value="1">
				<xsl:attribute name="onclick">setOrderNonPriceDirty();</xsl:attribute>
				<xsl:if test="ssExported=1">
				<xsl:attribute name="checked">checked</xsl:attribute>
				</xsl:if>
				</input>&amp;nbsp;<label for="ssExported">
				<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>This order has been exported
				</label>
		    </td>
		  </tr>
		  <tr>
		    <td>Internal Notes:</td>
		    <td>
		      <textarea name="ssInternalNotes" id="ssInternalNotes" rows="1" cols="60">
				<xsl:attribute name="title">These are the notes which only you can see. HTML is not supported for this field.</xsl:attribute>
				<xsl:attribute name="onfocus">this.rows=5;</xsl:attribute>
				<xsl:attribute name="onblur">this.rows=1;</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
		        <xsl:value-of select="ssInternalNotes" />
		      </textarea>
		    </td>
		  </tr>

		  <tr>
		    <td>External Notes:</td>
		    <td>
		      <textarea name="ssExternalNotes" id="ssExternalNotes" rows="1" cols="60">
				<xsl:attribute name="title">These are the notes which appear on the customer's order status page. This can be used in addition to the tracking number. HTML is not supported for this field.</xsl:attribute>
				<xsl:attribute name="onfocus">this.rows=5;</xsl:attribute>
				<xsl:attribute name="onblur">this.rows=1;</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
		        <xsl:value-of select="ssExternalNotes" />
		      </textarea>
		    </td>
		  </tr>

		  <tr>
		    <td>Sales Receipt Message:</td>
		    <td>
		      <textarea name="ssBackOrderMessage" id="ssBackOrderMessage" rows="1" cols="60">
				<xsl:attribute name="title">This appears on the sales receipt. HTML is not supported for this field.</xsl:attribute>
				<xsl:attribute name="onfocus">this.rows=5;</xsl:attribute>
				<xsl:attribute name="onblur">this.rows=1;</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
		        <xsl:value-of select="ssBackOrderMessage" />
		      </textarea>
		    </td>
		  </tr>

		  <tr>
		    <td>Packing Slip Message:</td>
		    <td>
		      <textarea name="ssBackOrderInternalMessage" id="ssBackOrderInternalMessage" rows="1" cols="60">
				<xsl:attribute name="title">This appears on the packing slip. HTML is not supported for this field.</xsl:attribute>
				<xsl:attribute name="onfocus">this.rows=5;</xsl:attribute>
				<xsl:attribute name="onblur">this.rows=1;</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
		        <xsl:value-of select="ssBackOrderInternalMessage" />
		      </textarea>
		    </td>
		  </tr>

		  <tr>
		    <td>Back Order Date Notified:</td>
		    <td>
				<input type="text" name="ssBackOrderDateNotified" id="ssBackOrderDateNotified">
					<xsl:attribute name="value"><xsl:value-of select="ssBackOrderDateNotified" /></xsl:attribute>
					<xsl:attribute name="ondblclick">this.value=this.form.todaysDate.value;this.onchange();</xsl:attribute>
					<xsl:attribute name="onchange">if (isDate(this,'Please enter a valid date.')){setOrderNonPriceDirty();}else{return false;}</xsl:attribute>
					<xsl:attribute name="title">Double click to set to today's date</xsl:attribute>
					<xsl:attribute name="size">20</xsl:attribute>
				</input>
				<a HREF="javascript:doNothing()" title="Select start date">
				<xsl:attribute name="onClick">setDateField(document.frmData.ssBackOrderDateNotified); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')</xsl:attribute>
				<img SRC="images/calendar.gif" BORDER="0" />
				</a>
		    </td>
		  </tr>

		  <tr>
		    <td>Shipping Email Sent On:</td>
		    <td>
				<input type="text" name="ssDateEmailSent" id="ssDateEmailSent">
					<xsl:attribute name="value"><xsl:value-of select="ssDateEmailSent" /></xsl:attribute>
					<xsl:attribute name="ondblclick">this.value=this.form.todaysDate.value;this.onchange();</xsl:attribute>
					<xsl:attribute name="onchange">if (isDate(this,'Please enter a valid date.')){setOrderNonPriceDirty();}else{return false;}</xsl:attribute>
					<xsl:attribute name="title">Double click to set to today's date</xsl:attribute>
					<xsl:attribute name="size">20</xsl:attribute>
				</input>
				<a HREF="javascript:doNothing()" title="Select start date">
				<xsl:attribute name="onClick">setDateField(document.frmData.ssDateEmailSent); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')</xsl:attribute>
				<img SRC="images/calendar.gif" BORDER="0" />
				</a>
		    </td>
		  </tr>

		  <xsl:if test="string-length($ordersExtra1_Label)>0">
		  <tr>
		    <td><xsl:value-of select="$ordersExtra1_Label" />:</td>
		    <td>
		      <input type="text" size="50" maxlength="50" name="ordersExtra1" id="ordersExtra1">
				<xsl:attribute name="title">HTML is not supported for this field.</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
				<xsl:attribute name="value"><xsl:value-of select="ordersExtra1" /></xsl:attribute>
		      </input>
		    </td>
		  </tr>
		  </xsl:if>

		  <tr>
		    <td>Edited by:</td>
		    <td>
		      <input type="text" size="50" maxlength="50" name="ssEditedBy" id="ssEditedBy">
				<xsl:attribute name="title">Person who edited this order.</xsl:attribute>
				<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
				<xsl:attribute name="value"><xsl:value-of select="ssEditedBy" /></xsl:attribute>
		      </input>
		    </td>
		  </tr>

		  <tr>
		    <td>Edited on:</td>
		    <td><xsl:value-of select="ssDateEdited" /></td>
		  </tr>

      </table>
    </td>
  </tr>
</xsl:template>

<xsl:template name="priorOrders">
	<xsl:if test="string-length(OrderHistory/PriorOrder/GroupName)>0">Customer is a member of Price Group: <xsl:value-of select="OrderHistory/PriorOrder/GroupName" /><br /></xsl:if>
	<fieldset onmouseover="showHideElement(document.all('tblPriorOrders'));" onmouseout="showHideElement(document.all('tblPriorOrders'));">
	<legend><xsl:value-of select="count(OrderHistory/PriorOrder)" /> prior order(s) totalling <xsl:value-of select="$currencySymbol" /><xsl:value-of select="OrderHistory/@priorOrderTotal" /></legend>
	<table class="tbl" style="border-collapse: collapse; border-color:#111111;display:none;" width="100%" cellpadding="0" cellspacing="0" border="1" id="tblPriorOrders">
		<colgroup>
		<col align="center" />
		<col align="center" />
		<col align="center" />
		<col align="center" />
		</colgroup>
		<tr class="tblhdr"><th>Order</th><th>Order Date</th><th>Amount</th><th>Items</th></tr>
	<xsl:for-each select="OrderHistory/PriorOrder">
		<tr>
		<td>
			<a>
				<xsl:attribute name="href">ssOrderAdmin.asp?OrderUID=<xsl:value-of select="uid" />&amp;Action=ViewOrder&amp;optDisplay=0&amp;optPayment_Filter=1&amp;optShipment_Filter=1</xsl:attribute>
				<xsl:value-of select="OrderNumber" />
			</a>
		</td>
		<td><xsl:value-of select="DateOrdered" /></td>
		<td><xsl:value-of select="$currencySymbol" /><xsl:value-of select="GrandTotal" /></td>
		<td><xsl:value-of select="SumOfQuantity" /></td>
		</tr>
	</xsl:for-each>
	</table>
	</fieldset>
</xsl:template>

<xsl:template name="shipmentStatus">

	<xsl:param name='shippingAddressCount' />

		  <tr class="tblhdr">
			<th colspan="2" align="left">Shipment Status</th>
		  </tr>
		  <tr>
		    <td colspan="2" align="left">
			  <xsl:for-each select="shippingAddress">
				<xsl:variable name="ShipMethod"><xsl:value-of select="ShipMethod" /></xsl:variable>
				<xsl:variable name="OrderAddressesUID"><xsl:value-of select="OrderAddressesUID" /></xsl:variable>
				<input type="hidden" name="OrderAddressesUID" id="OrderAddressesUID"><xsl:attribute name="value"><xsl:value-of select="$OrderAddressesUID" /></xsl:attribute></input>
				<input type="hidden">
					<xsl:attribute name="name">isDirty_OrderAddress_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
					<xsl:attribute name="id">isDirty_OrderAddress_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
					<xsl:attribute name="value">0</xsl:attribute>
				</input>
				<input type="hidden">
					<xsl:attribute name="name">ShipCarrierCode<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
					<xsl:attribute name="id">ShipCarrierCode<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
					<xsl:attribute name="value"><xsl:value-of select="ShipCarrierCode" /></xsl:attribute>
				</input>
		      <table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="1" cellpadding="2" cellspacing="0">
				<tr>
					<td colspan="5">
						<input type="checkbox" value="0">
						<xsl:attribute name="name">Pending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="id">Pending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="onclick">setOrderAddressDirty(<xsl:value-of select="$OrderAddressesUID" />);</xsl:attribute>
						<xsl:if test="Pending=0">
						<xsl:attribute name="checked">checked</xsl:attribute>
						</xsl:if>
						</input>
						&amp;nbsp;
						<label>
						<xsl:attribute name="for">Pending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>
						<xsl:text>This order has been shipped</xsl:text>
						</label>
					</td>
				</tr>
				<!--
				<tr>
					<td colspan="5">
						<input type="checkbox" value="0">
						<xsl:attribute name="name">BOPending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="id">BOPending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="onclick">setOrderAddressDirty(<xsl:value-of select="$OrderAddressesUID" />);</xsl:attribute>
						<xsl:if test="BOPending=0">
						<xsl:attribute name="checked">checked</xsl:attribute>
						</xsl:if>
						</input>
						&amp;nbsp;
						<label>
						<xsl:attribute name="for">BOPending<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
						<xsl:attribute name="onclick">labelClick(this);</xsl:attribute>
						<xsl:text>This order has been shipped (Backorder)</xsl:text>
						</label>
					</td>
				</tr>
				-->
				<tr>
					<td colspan="5">Ship via:&amp;nbsp;
						<select>
							<xsl:attribute name="name">ShipMethod<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">ShipMethod<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="onchange">shipMethodChanged(this, <xsl:value-of select="$OrderAddressesUID" />);</xsl:attribute>
							<xsl:for-each select="../../ShippingCarriers/ShippingCarrier">
							<option>
							<xsl:attribute name="value"><xsl:value-of select="@ShippingCode" /></xsl:attribute>
							<xsl:if test="@Method = $ShipMethod"><xsl:attribute name="selected" /></xsl:if>
							<xsl:value-of select="@Method" />
							</option>
							</xsl:for-each>
						</select>
					</td>
				</tr>
		        <tr>
		          <xsl:if test="$shippingAddressCount > 1">
		          <th>Item</th>
		          <th>Shipping Address</th>
		          </xsl:if>
		          <th>Tracking Number</th>
		          <th>Message</th>
		          <th>Ship Date</th>
		          <xsl:if test="string-length($PackageWeight_Label + $Insured_Label + $orderTrackingExtra1_Label)>0">
		          <th><xsl:value-of select="$PackageWeight_Label" /></th>
		          </xsl:if>
		        </tr>

					<xsl:for-each select="odrdtOrderTracking">
						<xsl:variable name="OrderTrackingUID"><xsl:value-of select="OrderTrackingUID" /></xsl:variable>
						<input type="hidden" name="OrderTrackingUID"><xsl:attribute name="value"><xsl:value-of select="$OrderTrackingUID" /></xsl:attribute></input>
						<xsl:variable name="EvenRow" select="position() mod 2"/>
						<tr>
							<xsl:if test="$EvenRow = 0">
								<xsl:attribute name="bgcolor">lightgrey</xsl:attribute>
							</xsl:if>

							<xsl:if test="$shippingAddressCount > 1">
								<td valign="top">
									<xsl:for-each select="key('orderItemShippingKey',$OrderAddressesUID)">
										<xsl:value-of select="ProductCode" /> - <xsl:value-of select="ProductName" />(Qty:<xsl:value-of select="Quantity" />)<br />
									</xsl:for-each>
								</td>
								<td valign="top">
									<table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="0" cellpadding="2" cellspacing="0">
										<colgroup align="left" />
										<tr>
											<th align="left">
											<xsl:value-of select="../NickName" />
											<xsl:if test="BackOrderFlag = 1">(Backorder)</xsl:if>
											</th>
										</tr>
										<tr>
											<td>
											<a href="" title="Click to Edit Billing Information">
											<xsl:attribute name="onclick">OpenHelp('ssOrderAdmin_Customer.asp?uid=<xsl:value-of select="../uid" />'); return false;</xsl:attribute><xsl:call-template name="outputName"><xsl:with-param name='currentNode' select='../.' /></xsl:call-template></a>
											</td>
										</tr>
										<xsl:if test="string-length(../Company)>0">
											<tr><td><xsl:value-of select="../Company" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(../Address1)>0">
											<tr><td><xsl:value-of select="../Address1" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(Address2)>0">
											<tr><td><xsl:value-of select="../Address2" /></td></tr>
										</xsl:if>
										<tr>
											<td>
											<xsl:value-of select="../City"/>, <xsl:value-of select="../State"/><xsl:text> </xsl:text><xsl:value-of select="../Zip"/>
											</td>
										</tr>
										<xsl:if test="../Country[.!='US']">
											<tr><td><xsl:value-of select="../CountryName" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(../Phone)>0">
											<tr><td><xsl:value-of select="../Phone" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(../Fax)>0">
											<tr><td><xsl:value-of select="../Fax" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(../EMail)>0">
											<tr><td><xsl:value-of select="../EMail" /></td></tr>
										</xsl:if>
										<xsl:if test="string-length(../SpecialInstruction)>0">
											<tr><td>Special Instructions:<br /><xsl:value-of select="../SpecialInstruction" /></td></tr>
										</xsl:if>
										</table>
									</td>
								</xsl:if><!-- test="$shippingAddressCount > 1" -->

							<td valign="top">
							<input type="hidden" name="OrderTrackingOrderAddressesUID"><xsl:attribute name="value"><xsl:value-of select="ShipToAddressID" /></xsl:attribute></input>
							<input type="checkbox" name="deleteTrackingNumber" id="deleteTrackingNumber" title="Delete this tracking item">
								<xsl:attribute name="value"><xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
							</input>
							<input type="text">
								<xsl:attribute name="name">TrackingNumber_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="id">TrackingNumber_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="TrackingNumber" /></xsl:attribute>
								<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
							</input>
							</td>
							<td valign="top">
							<textarea rows="2" cols="10">
								<xsl:attribute name="name">TrackingMessage_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="id">TrackingMessage_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="title">This is the message which appears on the customer's package tracking screen. Double-click this field to use the HTML edit window.</xsl:attribute>
								<xsl:attribute name="onfocus">this.cols=40;this.rows=5;</xsl:attribute>
								<xsl:attribute name="onblur">this.cols=10;this.rows=2;</xsl:attribute>
								<xsl:attribute name="ondblclick">return openACE(this);</xsl:attribute>
								<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
							<xsl:value-of select="TrackingMessage" />
							</textarea>
							</td>
							<td valign="top">
							<input type="text">
								<xsl:attribute name="name">SentDate_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="id">SentDate_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="SentDate" /></xsl:attribute>
								<xsl:attribute name="ondblclick">this.value=this.form.todaysDate.value;this.onchange();</xsl:attribute>
								<xsl:attribute name="onchange">if (isDate(this,'Please enter a valid date.')){setOrderTrackingDirty();}else{return false;}</xsl:attribute>
								<xsl:attribute name="title">Double click to set date payment received to today's date</xsl:attribute>
								<xsl:attribute name="size">10</xsl:attribute>
							</input>
							<a HREF="javascript:doNothing()" title="Select start date">
							<xsl:attribute name="onClick">setDateField(document.frmData.SentDate_<xsl:value-of select="$OrderTrackingUID" />); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')</xsl:attribute>
							<img SRC="images/calendar.gif" BORDER="0" />
							</a>
							</td>
							<xsl:if test="string-length($PackageWeight_Label + $Insured_Label + $orderTrackingExtra1_Label)>0">
							<td valign="top">
							<xsl:if test="string-length($PackageWeight_Label)>0">
							<input type="text" size="4">
								<xsl:attribute name="name">PackageWeight_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="id">PackageWeight_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="title"><xsl:value-of select="$PackageWeight_Label" /></xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="PackageWeight" /></xsl:attribute>
								<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
							</input>
							</xsl:if>
							
							<xsl:if test="string-length($Insured_Label)>0">
							<input type="checkbox" value="1">
								<xsl:attribute name="name">Insured_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="id">Insured_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
								<xsl:attribute name="title"><xsl:value-of select="$Insured_Label" /></xsl:attribute>
								<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
								<xsl:if test="Insured=1">
								<xsl:attribute name="checked">checked</xsl:attribute>
								</xsl:if>
							</input>
							</xsl:if>

							<xsl:if test="string-length($orderTrackingExtra1_Label)>0">
								<br />
								<input type="text" size="8" maxlength="50">
									<xsl:attribute name="name">orderTrackingExtra1_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
									<xsl:attribute name="id">orderTrackingExtra1_<xsl:value-of select="$OrderTrackingUID" /></xsl:attribute>
									<xsl:attribute name="title"><xsl:value-of select="$orderTrackingExtra1_Label" /></xsl:attribute>
									<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
									<xsl:attribute name="value"><xsl:value-of select="orderTrackingExtra1" /></xsl:attribute>
									<xsl:attribute name="onfocus">this.size=50;</xsl:attribute>
									<xsl:attribute name="onblur">this.size=8;</xsl:attribute>
								</input>
							</xsl:if>
							</td>
							</xsl:if>
						</tr>
					</xsl:for-each><!-- select="odrdtOrderTracking" -->
					
					<!-- added here to add additional tracking numbers -->
					<tr>
						<xsl:if test="$shippingAddressCount > 1">
							<td valign="top" colspan="2">&amp;nbsp;</td>
						</xsl:if><!-- test="$shippingAddressCount > 1" -->
						<td valign="top" align="right">
						<input type="hidden" name="OrderTrackingOrderAddressesUID"><xsl:attribute name="value"><xsl:value-of select="OrderAddressesUID" /></xsl:attribute></input>
						<input type="hidden" name="OrderTrackingUID" id="OrderTrackingUID">
							<xsl:attribute name="value">NONE_<xsl:value-of select="OrderAddressesUID" /></xsl:attribute>
						</input>
						<input type="text" title="Add an additional tracking number for this package">
							<xsl:attribute name="name">TrackingNumber_NONE_<xsl:value-of select="OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">TrackingNumber_NONE_<xsl:value-of select="OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="value"><xsl:value-of select="TrackingNumber" /></xsl:attribute>
							<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
						</input>
						</td>
						<td valign="top">
						<textarea rows="2" cols="10">
							<xsl:attribute name="name">TrackingMessage_NONE_<xsl:value-of select="OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">TrackingMessage_NONE_<xsl:value-of select="OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="title">This is the message which appears on the customer's package tracking screen. Double-click this field to use the HTML edit window.</xsl:attribute>
							<xsl:attribute name="onfocus">this.cols=40;this.rows=5;</xsl:attribute>
							<xsl:attribute name="onblur">this.cols=10;this.rows=2;</xsl:attribute>
							<xsl:attribute name="ondblclick">return openACE(this);</xsl:attribute>
							<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
						<xsl:value-of select="TrackingMessage" />
						</textarea>
						</td>
						<td valign="top">
						<input type="text" name="SentDate" id="SentDate">
							<xsl:attribute name="name">SentDate_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">SentDate_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="value"><xsl:value-of select="SentDate" /></xsl:attribute>
							<xsl:attribute name="ondblclick">this.value=this.form.todaysDate.value;this.onchange();</xsl:attribute>
							<xsl:attribute name="onchange">if (isDate(this,'Please enter a valid date.')){setOrderTrackingDirty();}else{return false;}</xsl:attribute>
							<xsl:attribute name="title">Double click to set date payment received to today's date</xsl:attribute>
							<xsl:attribute name="size">10</xsl:attribute>
						</input>
						<a HREF="javascript:doNothing()" title="Select start date">
						<xsl:attribute name="onClick">setDateField(document.frmData.SentDate_NONE_<xsl:value-of select="$OrderAddressesUID" />); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')</xsl:attribute>
						<img SRC="images/calendar.gif" BORDER="0" />
						</a>
						</td>

						<xsl:if test="string-length($PackageWeight_Label + $Insured_Label + $orderTrackingExtra1_Label)>0">
						<td valign="top">
						<xsl:if test="string-length($PackageWeight_Label)>0">
						<input type="text" size="4">
							<xsl:attribute name="name">PackageWeight_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">PackageWeight_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="title"><xsl:value-of select="$PackageWeight_Label" /></xsl:attribute>
							<xsl:attribute name="value"><xsl:value-of select="PackageWeight" /></xsl:attribute>
							<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
						</input>
						</xsl:if>
						
						<xsl:if test="string-length($Insured_Label)>0">
						<input type="checkbox" value="1">
							<xsl:attribute name="name">Insured_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="id">Insured_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
							<xsl:attribute name="title"><xsl:value-of select="$Insured_Label" /></xsl:attribute>
							<xsl:attribute name="onchange">setOrderTrackingDirty();</xsl:attribute>
							<xsl:if test="Insured=1">
							<xsl:attribute name="checked">checked</xsl:attribute>
							</xsl:if>
						</input>
						</xsl:if>

						<xsl:if test="string-length($orderTrackingExtra1_Label)>0">
							<br />
							<input type="text" size="8" maxlength="50">
								<xsl:attribute name="name">orderTrackingExtra1_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
								<xsl:attribute name="id">orderTrackingExtra1_NONE_<xsl:value-of select="$OrderAddressesUID" /></xsl:attribute>
								<xsl:attribute name="title"><xsl:value-of select="$orderTrackingExtra1_Label" /></xsl:attribute>
								<xsl:attribute name="onchange">setOrderNonPriceDirty();</xsl:attribute>
								<xsl:attribute name="value"><xsl:value-of select="orderTrackingExtra1" /></xsl:attribute>
								<xsl:attribute name="onfocus">this.size=50;</xsl:attribute>
								<xsl:attribute name="onblur">this.size=8;</xsl:attribute>
							</input>
						</xsl:if>
						</td>
						</xsl:if>

					</tr>

					<xsl:if test="string-length(SpecialInstruction)>0">
						<tr>
							<td>Special Instructions:</td>
							<td colspan="4">
							<div>
								<xsl:attribute name="title">Double-click this entry to change the special instructions.</xsl:attribute>
								<xsl:attribute name="ondblclick">OpenHelp('ssOrderAdmin_Customer.asp?uid=<xsl:value-of select="uid" />'); return false;</xsl:attribute>
								<xsl:value-of select="SpecialInstruction" />
							</div>
							</td>
						</tr>
					</xsl:if>
		      </table>
			  </xsl:for-each><!-- select="shippingAddress" -->
		    </td>
		  </tr>

</xsl:template>

<xsl:template name="orderDetailFooter">

<div id="divEmail" style="position:absolute; display:none">
<table class="tbl" style="border-style:outset; border-color:steelblue;" border="3" cellspacing="0" cellpadding="0" bgcolor="white" id="tblEmail">
<tr><td>
<table class="tbl" style="border-collapse: collapse; border-color:steelblue;" border="0" width="100%" cellspacing="0" cellpadding="3">
  <input type="hidden" id="StockEmail" name="StockEmail" value="1" />
  <tr class="tblhdr">
    <th>&amp;nbsp;</th>
    <th align="left">Select an email template</th>
  </tr>
  <tr>
    <td align="right">&amp;nbsp;</td>
    <td><br />
      <script language="javascript">
      function changeEmailTemplate(Index)
      {
      frmData.emailSubject.value  = document.all("enEmail_Subject" + Index).value;
      frmData.emailBody.value  = document.all("enEmail_Body" + Index).value;
      }
      
      function editEmail_click(theItem)
      {
		document.all("divEmail").style.display = "";
		ReplaceEmailText();
		document.frmData.StockEmail.value=0;
		document.frmData.emailBody.focus();
		ScrollToElem("btnSendEmail");
		return false;      
      }
      </script>
      <select name="emailFile" ID="emailFile" onchange="changeEmailTemplate(this.selectedIndex); return false;">
      </select>
    </td>
  </tr>
  <tr>
    <td align="right">From:</td>
    <td><input type="text" name="emailFrom" ID="emailFrom" size="75" VALUE="" /></td>
  </tr>
  <tr>
    <td align="right">To:</td>
    <td><input type="text" name="emailTo" ID="emailTo" size="75" VALUE="" /></td>
  </tr>
  <tr>
    <td align="right">Subject:</td>
    <td><input type="text" name="emailSubject" ID="emailSubject" size="75" VALUE="" /></td>
  </tr>
  <tr>
    <td align="right">Body:</td>
    <td><textarea name="emailBody" ID="emailBody" rows="12" cols="70"></textarea></td>
  </tr>
  <tr>
    <td>&amp;nbsp;</td>
    <td>
        <input class="butn" type="button" value="Send" name="btnSendEmail" ID="btnSendEmail" onclick="frmData.SendEmail.checked=true; frmData.SendEmail.onclick(); document.all('divEmail').style.display = 'none';" />&amp;nbsp;
        <input class="butn" type="button" value="Cancel" name="Cancel" ID="Cancel" onclick='document.all("divEmail").style.display = "none";' />
    </td>
  </tr>
</table>
</td></tr>
</table>
</div>

&amp;nbsp;<input type="checkbox" name="SendEmail" id="SendEmail" value="1" onclick="enableSave();" title="You must 'Save Changes' for the email to send." />&amp;nbsp;<label for="SendEmail" onclick="frmData.SendEmail.onclick();">Send Email</label>&amp;nbsp;<a href="" onclick='editEmail_click(this); return false;' title='Customize the email'>(Edit)</a>&amp;nbsp;|&amp;nbsp;<a href="" onclick="OpenHelp('ssOrderAdmin_EmailConfigure.asp'); return false;">(Configure Templates)</a><br />
 
<table class="tbl" style="border-collapse: collapse; border-color:steelblue;" border="0" width="100%" cellspacing="0" cellpadding="3">
	<tr>
		<td>
		&amp;nbsp;<a href="" title="View the packing slip for this order" onclick="viewOrder_Special(cstrSalesReceiptTemplate); return false;">Sales Receipt</a>
		&amp;nbsp;|&amp;nbsp;<a href="" title="View the packing slip for this order" onclick="viewOrder_Special(cstrPackingSlipTemplate); return false;">Packing Slip</a>
		&amp;nbsp;|&amp;nbsp;<a href="" title="View the checklist for this order" onclick="viewOrder_Special(cstrCheckListTemplate); return false;">Checklist</a>
		</td>
	</tr>
	
	<tr><td>&amp;nbsp;<a title="View the status page in merchant tools for this order" target="merchantTools"><xsl:attribute name="href">../OrderStatus.aspx?OrderId=<xsl:value-of select="OrderNumber" /></xsl:attribute>Status Page - Merchant tools</a></td></tr>
	<tr>
		<td>&amp;nbsp;<a title="View the order details in merchant tools for this order" target="merchantTools"><xsl:attribute name="href">../orddetails.aspx?OrderID=<xsl:value-of select="OrderNumber" /></xsl:attribute>Order Details - Merchant tools</a></td>
		<td></td><td rowspan="4" valign="bottom" align="right"><a class="copyright" onclick="return false;" title=""><xsl:value-of select="$styleSheetVersion" /></a></td>
	</tr>
</table>
</xsl:template>

</xsl:stylesheet>