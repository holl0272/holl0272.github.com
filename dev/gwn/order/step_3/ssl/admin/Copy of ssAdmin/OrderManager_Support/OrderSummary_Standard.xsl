<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:import href="../exportTemplates/supportingTemplates/supportTemplates.xsl"/>
<xsl:import href="OrderDetails_Common.xsl"/>
<xsl:import href="OrderDetails_Configuration.xsl"/>
<xsl:output method="html" />
<!--

Release Version:	1.01.008
Revision Date:		January 18, 2005

-->
<xsl:variable name="styleSheetVersion" select="'Version 1.01.008'"/>

  <xsl:template match='/'>
<table class="tbl"  width="100%" cellpadding="3" cellspacing="0" border="0" bgcolor="whitesmoke" id="tblSummary">
  <colgroup align="left" />
  <colgroup align="left" />
  <tr>
    <td width="100%" colspan="2">
	<xsl:apply-templates select="orders" />
    </td>
  </tr>
</table>
  </xsl:template>

  <xsl:template match="orders" >
      <xsl:variable name="BackOrderCount" select="ItemsOnBackOrder"/>
      <xsl:variable name="displayBackOrder" select="'true'"/>
      <xsl:variable name="SortOrder">
		<xsl:if test="SortOrder = 'ASC'">DESC</xsl:if>
		<xsl:if test="SortOrder = 'DESC'">ASC</xsl:if>
      </xsl:variable>

      <table class="tbl" style="border-collapse: collapse; border-color:#111111;" border="0" cellspacing="0" cellpadding="1" width="100%" id="tblOrderDetailSummary">
		<colgroup width="1%" align="center" valign="top" />
		<colgroup width="1%" align="center" valign="top" />
		<colgroup width="8%" align="center" valign="top" />
		<colgroup width="40%" align="left" valign="top" />
		<colgroup width="10%" align="center" valign="top" />
		<colgroup width="10%" align="center" valign="top" />
		<colgroup width="10%" align="center" valign="top" />
		<colgroup width="10%" align="center" valign="top" />
		<colgroup width="10%" align="center" valign="top" />

		<tr class="tblhdr">
			<th valign="middle">&amp;nbsp;<input type="checkbox" name="chkCheckAll" id="chkCheckAll1"  onclick="checkAll(theDataForm.chkOrderUID, this.checked); checkAll(theDataForm.chkCheckAll2, this.checked);" value="" /></th>
			<th valign="middle" style="cursor:hand;" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by flagged orders">
				<xsl:attribute name="onclick">SortColumn('ssOrderFlagged','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>&amp;nbsp;</xsl:text>
				<xsl:if test="OrderBy = 'ssOrderFlagged'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(2,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Order Numbers">
				<xsl:attribute name="onclick">SortColumn('OrderNumber','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Order Number</xsl:text>
				<xsl:if test="OrderBy = 'OrderNumber'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(3,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Last Names">
				<xsl:attribute name="onclick">SortColumn('LastName','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Last Name</xsl:text>
				<xsl:if test="OrderBy = 'LastName'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(4,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Item Quantities">
				<xsl:attribute name="onclick">SortColumn('Quantity','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Items</xsl:text>
				<xsl:if test="OrderBy = 'Quantity'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(5,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Order Totals">
				<xsl:attribute name="onclick">SortColumn('GrandTotal','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Order Total</xsl:text>
				<xsl:if test="OrderBy = 'GrandTotal'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(6,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Order Date">
				<xsl:attribute name="onclick">SortColumn('DateOrdered','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Order Date</xsl:text>
				<xsl:if test="OrderBy = 'DateOrdered'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onclick="SortColumn(7,'DESC');" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Payment Received Dates">
				<xsl:attribute name="onclick">SortColumn('PaymentsPending','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Payment Received</xsl:text>
				<xsl:if test="OrderBy = 'PaymentsPending'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
			<th valign="middle" style="cursor:hand;" onMouseOver="HighlightColor(this); return DisplayTitle(this);" onMouseOut="deHighlightColor(this); ClearTitle();" title="Sort by Order Shipped Dates">
				<xsl:attribute name="onclick">SortColumn('SentDate','<xsl:value-of select="$SortOrder" />');</xsl:attribute>
				<xsl:text>Order Shipped</xsl:text>
				<xsl:if test="OrderBy = 'SentDate'">
					<xsl:text>&amp;nbsp;</xsl:text>
					<xsl:if test="SortOrder = 'ASC'"><img src="images/up.gif" border="0" align="bottom" /></xsl:if>
					<xsl:if test="SortOrder = 'DESC'"><img src="images/down.gif" border="0" align="bottom" /></xsl:if>
				</xsl:if>
			</th>
		</tr>
	
		<xsl:for-each select="order">
			<xsl:variable name="EvenRow" select="position() mod 2"/>
			<tr>
				<xsl:attribute name="class"><xsl:value-of select="TRClass" /></xsl:attribute>
				<xsl:attribute name="title">Click to view order number <xsl:value-of select="OrderNumber" /></xsl:attribute>
				
				<xsl:attribute name="onmouseover">doMouseOverRow(this); DisplayTitle(this);</xsl:attribute>
				<xsl:attribute name="onmouseout">doMouseOutRow(this); ClearTitle();</xsl:attribute>

				<td>
					<input type="checkbox" name="chkOrderUID" id="chkOrderUID">
						<xsl:attribute name="value"><xsl:value-of select="@uid" /></xsl:attribute>
						<xsl:if test="Checked='True'">
						<xsl:attribute name="checked">checked</xsl:attribute>
						</xsl:if>
						<xsl:if test="TRImage!=''">
							<img border="0">
								<xsl:attribute name="src"><xsl:value-of select="TRImage" /></xsl:attribute>
							</img>
						</xsl:if>
					</input>
				</td>
				<xsl:if test="ssOrderFlagged = '1'">
					<td><img src="images/MSGBOX03.ICO" alt="x" title="flagged for follow up" height="12" /></td>
				</xsl:if>
				<xsl:if test="ssOrderFlagged != '1'">
					<td>&amp;nbsp;</td>
				</xsl:if>
				<td>
				<xsl:attribute name="onmousedown">ViewOrder('<xsl:value-of select="@uid" />');</xsl:attribute>
				<u><xsl:value-of select="OrderNumber" /></u>
				</td>
				<td><xsl:value-of select="LastName" /></td>
				<td><xsl:value-of select="SumOfQuantity" /></td>
				<td>
					<xsl:value-of select="$currencySymbol" />
					<xsl:call-template name="customNumber"><xsl:with-param name="currentNode" select="string(GrandTotal)" /></xsl:call-template>
				</td>
				<td><xsl:value-of select="DateOrdered" /></td>
				<td><xsl:value-of select="PaymentsPending" /></td>
				<td><xsl:value-of select="SentDate" /></td>
			</tr>

		</xsl:for-each>

		<script language="javascript">
			var aryItemList = new Array();

			tipMessage['prevItem']=["View previous order", "Order Info"]
			tipMessage['nextItem']=["View next order", "Order Info"]
			tipMessage['currentItem']=["Current order", "Order Info"]

			function initializeNextItemControls()
			{
				if (aryItemList[0][0] != 0)
				{
					showElement(document.all("prevItem"));
					if (aryItemList[0][0] == -1)
					{
						tipMessage['prevItem'][0] = "View previous page of orders";
						tipMessage['prevItem'][1] = "";
					}else{
						tipMessage['prevItem'][1] = aryItemList[0][1];
					}
				}

				if (aryItemList[2][0] != 0)
				{
					showElement(document.all("nextItem"));
					if (aryItemList[2][0] == -1)
					{
						tipMessage['nextItem'][0] = "View next page of orders";
						tipMessage['nextItem'][1] = "";
					}else{
						tipMessage['nextItem'][1] = aryItemList[2][1];
					}
				}

				if (aryItemList[1][0] != 0)
				{
					tipMessage['currentItem'][1] = aryItemList[1][1];
					document.all("orderDetailNumber").innerText = " " + aryItemList[1][1] + " ";
				}else{
					tipMessage['currentItem'][1] = "";
				}
			}

			function selectPrevItem()
			{
				if (aryItemList[0][0] == -1)
				{
					var num = new Number(theDataForm.AbsolutePage.value);
					ViewPage(num - 1);
				}else{
					ViewOrder(aryItemList[0][0]);
				}
				return false;
			}

			function selectNextItem()
			{
				if (aryItemList[2][0] == -1)
				{
					var num = new Number(theDataForm.AbsolutePage.value);
					ViewPage(num + 1);
				}else{
					ViewOrder(aryItemList[2][0]);
				}
				return false;
			}

			aryItemList[0] = new Array(0, "");
			aryItemList[1] = new Array(0, "");
			aryItemList[2] = new Array(0, "");
			
			<xsl:for-each select="order">
				<xsl:if test="ActiveOrder = '1'">
					<xsl:variable name="preceding" select="preceding-sibling::order[1]"/>
					<xsl:variable name="current" select="."/>
					<xsl:variable name="following" select="following-sibling::order"/>
					
					//PageCount: <xsl:value-of select="../PageCount" />
					//AbsolutePage: <xsl:value-of select="../AbsolutePage" />

					if ('<xsl:value-of select="$preceding/@uid" />' == '')
					{
						<xsl:if test="../AbsolutePage &gt; 1">aryItemList[0] = new Array(-1, "", "");</xsl:if>
					}else{
						aryItemList[0] = new Array("<xsl:value-of select="$preceding/@uid" />", "<xsl:value-of select="$preceding/OrderNumber" />: <xsl:value-of select="$preceding/LastName" /><br /><xsl:value-of select="$preceding/SumOfQuantity" /> item(s)<br /><xsl:value-of select="$currencySymbol" /><xsl:call-template name="customNumber"><xsl:with-param name="currentNode" select="string($preceding/GrandTotal)" /></xsl:call-template>");
					}
					
					if ('<xsl:value-of select="$current/@uid" />' != '')
					{
						aryItemList[1] = new Array("<xsl:value-of select="$current/@uid" />", "<xsl:value-of select="$current/OrderNumber" />");
					}
					
					if ('<xsl:value-of select="$following/@uid" />' == '')
					{
						<xsl:if test="../PageCount &gt; ../AbsolutePage">aryItemList[2] = new Array(-1, "", "");</xsl:if>
					}else{
						aryItemList[2] = new Array("<xsl:value-of select="$following/@uid" />", "<xsl:value-of select="$following/OrderNumber" />: <xsl:value-of select="$following/LastName" /><br /><xsl:value-of select="$following/SumOfQuantity" /> item(s)<br /><xsl:value-of select="$currencySymbol" /><xsl:call-template name="customNumber"><xsl:with-param name="currentNode" select="string($following/GrandTotal)" /></xsl:call-template>");
					}

				</xsl:if>
			</xsl:for-each>
			
			initializeNextItemControls();
		</script>

		<tr><td colspan="9" align="center"><hr width="100%" /></td></tr>	

		<tr>
			<td><input type="checkbox" name="chkCheckAll" id="chkCheckAll2"  onclick="checkAll(theDataForm.chkOrderUID, this.checked); checkAll(theDataForm.chkCheckAll1, this.checked);" value="" /></td>
					<!--<u><xsl:value-of select="sum(OrderItems/Quantity)" /> item<xsl:if test="sum(OrderItems/Quantity) > 1">s</xsl:if></u><br />
					<u><xsl:value-of select="count(OrderItems/Quantity)" /> items(s)</u><br />-->
			<td>&amp;nbsp;</td>
			<td colspan="2" align="left"><xsl:value-of select="count(order)" /> order(s)</td>
			<td align="center"><xsl:value-of select="sum(order/SumOfQuantity)" /> item(s)</td>
			<td>&amp;nbsp;</td>
			<td>
				<xsl:value-of select="$currencySymbol" />
				<xsl:call-template name="customNumber"><xsl:with-param name="currentNode" select="string(sum(order/GrandTotal))" /></xsl:call-template>
			</td>
			<td colspan="4">&amp;nbsp;</td>
		</tr>

		<tr class="tblhdr">
			<th colspan="9" align="center">
				<xsl:value-of select="RecordCount" /> Orders match your search criteria<br />
				<xsl:text>Show </xsl:text>
				<input type="text" name="PageSize" id="PageSize" maxlength="4" size="4" style="text-align:center;">
					<xsl:attribute name="onblur">return isInteger(this, true, 'Please enter a positive integer for the recordset page size.');</xsl:attribute>
					<xsl:attribute name="ondblclick">this.value=<xsl:value-of select="RecordCount" />;</xsl:attribute>
					<xsl:attribute name="value"><xsl:value-of select="MaxRecords" /></xsl:attribute>
				</input>
				<xsl:text>&amp;nbsp;</xsl:text>
				<a href="" class="tblhdr" onclick="document.frmData.submit(); return false;" title="Set records to show">orders</a> at a time.

				<xsl:if test="count(orderPaging/Page) > 1">

					<xsl:if test="AbsolutePage != 1">
						<a href="#">
						<xsl:attribute name="onclick">return ViewPage(<xsl:value-of select="number(AbsolutePage)-1" />);</xsl:attribute>
						<xsl:text>&lt;&lt;</xsl:text>
						</a><xsl:text>&amp;nbsp;</xsl:text>
					</xsl:if>

					<xsl:for-each select="orderPaging/Page">
						<xsl:if test="../../AbsolutePage = ."><xsl:value-of select="." /></xsl:if>
						<xsl:if test="../../AbsolutePage != .">
							<a href="#">
								<xsl:attribute name="onclick">return ViewPage(<xsl:value-of select="." />);</xsl:attribute>
								<xsl:value-of select="." />
							</a>
						</xsl:if>
						<xsl:text>&amp;nbsp;</xsl:text>
					</xsl:for-each>

					<xsl:if test="AbsolutePage != PageCount">
						<a href="#">
						<xsl:attribute name="onclick">return ViewPage(<xsl:value-of select="number(AbsolutePage)+1" />);</xsl:attribute>
						<xsl:text>&gt;&gt;</xsl:text>
						</a>
					</xsl:if>

				</xsl:if>

			</th>
		</tr>
      </table>
  </xsl:template>

<xsl:template name="outputName">
	<xsl:param name="currentNode" />
	<xsl:value-of select="$currentNode/FirstName"/><xsl:text> </xsl:text>
	<xsl:if test="string-length($currentNode/MI)>0"><xsl:value-of select="$currentNode/MI" /><xsl:text> </xsl:text></xsl:if>
	<xsl:value-of select="$currentNode/LastName" />
</xsl:template>

</xsl:stylesheet>