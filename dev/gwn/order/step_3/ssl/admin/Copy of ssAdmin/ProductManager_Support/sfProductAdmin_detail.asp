<% Sub WriteProductDetail(byVal strProductID) %>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblitemDetail">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
  <tr class="tblhdr">
	<th align=center><span id="spanprodName"></span>&nbsp;</th>
  </tr>
  <tr>
    <td>
	<% If mblnShowTabs Then %>
	<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" ID="Table2">
		<tr class="tblhdr" align=center>
			<td nowrap ID="tdGeneral" class="hdrNonSelected" onclick="return DisplaySection('General');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View General Product Information">General</td>
			<td ID='tdSpacer1' bgcolor="white">&nbsp;</td>
			<td nowrap ID="tdDetail" class="hdrNonSelected" onclick="return DisplaySection('Detail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Details" >Detail</td>
			<td ID='tdSpacer2' bgcolor="white">&nbsp;</td>
			<td nowrap ID="tdAttributes" class="hdrNonSelected" onclick='return DisplaySection("Attributes");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Attributes" >Attributes</td>
			<td ID='tdSpacer3' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdShipping' class="hdrNonSelected" onclick='return DisplaySection("Shipping");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Shipping Settings" >Shipping</td>
		<% If cblnSF5AE Then %>
			<td ID='tdSpacer4' bgcolor="white">&nbsp;</td>
			<td nowrap ID="tdMTP" class="hdrNonSelected" onclick='return DisplaySection("MTP");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Multi-Tier Pricing" >Multi-Tier Pricing</td>
			<td ID='tdSpacer5' bgcolor="white">&nbsp;</td>
			<td nowrap ID="tdInventory" class="hdrNonSelected" onclick='return DisplaySection("Inventory");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Inventory" >Inventory</td>
			<td ID='tdSpacer6' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdCategory' class="hdrNonSelected" onclick='return DisplaySection("Category");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Category Settings" >Category</td>
		<% End If %>
		<% If mclsProduct.CustomMTP Then %>
			<td ID='tdSpacer7' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdMTP' class="hdrNonSelected" onclick='return DisplaySection("MTP");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Multi-Tier Settings" >Multi-Tier Pricing</td>
		<% End If 'cblnSF5AE %>
		<% If cblnAddon_DynamicProductDisplay Then %>
			<td ID='tdSpacer9' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdRelatedProducts' class="hdrNonSelected" onclick='return DisplaySection("RelatedProducts");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Related Products">Related Products</td>
		<% End If 'cblnAddon_DynamicProductDisplay %>
		<% If cblnAddon_ProductReview Then %>
			<td ID='tdSpacer19' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdProductReview' class="hdrNonSelected" onclick='return DisplaySection("ProductReview");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Reviews">Reviews</td>
		<% End If 'cblnAddon_ProductReview %>
		<% If isArray(mclsProduct.CustomValues) Then %>
			<td ID='tdSpacer8' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdCustom' class="hdrNonSelected" onclick='return DisplaySection("Custom");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Custom Settings" >Custom</td>
		<% End If 'cblnSF5AE %>
			<td ID='tdSpacer20' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdProductSales' class="hdrNonSelected" onclick='return DisplaySection("ProductSales");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View product sales" >Sales</td>
			<td ID='tdSpacer21' bgcolor="white">&nbsp;</td>
			<td nowrap ID='tdSEO' class="hdrNonSelected" onclick='return DisplaySection("SEO");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View SEO" >SEO</td>
			<td ID='tdEmpty' bgcolor="white" width="100%">&nbsp;</td>
		</tr>
		<tr class="tblhdr" align=center>
		<%
		Dim plngNumCells
		
		plngNumCells = 8	'base count
		If cblnSF5AE Then plngNumCells = plngNumCells + 6
		If cblnAddon_DynamicProductDisplay Then plngNumCells = plngNumCells + 2
		If cblnAddon_ProductReview Then plngNumCells = plngNumCells + 2
		If mclsProduct.CustomMTP Then plngNumCells = plngNumCells + 2
		If isArray(mclsProduct.CustomValues) Then plngNumCells = plngNumCells + 2
		plngNumCells = plngNumCells + 4	'for sales, SEO
		%>
		<td colspan="<%= plngNumCells %>" height="8pt"></td>
		</tr>
	</table>
	<% Else %>
	<table width="100%" border="0" rules="none" ID="Table3"><tr><td></td></tr></table>
	<% End If	'mblnShowTabs %>

	<% 
	Call WriteGeneralTable
	Call WriteDetailTable
	Call WriteShippingTable
	Call WriteAttributeTable
	If cblnSF5AE Then Call WriteMTPTable
	If mclsProduct.CustomMTP Then Call WriteCustomMTPTable
	If cblnSF5AE Then Call WriteCategoryTable
	If cblnSF5AE Then Call WriteInventoryTable
	If cblnAddon_DynamicProductDisplay Then Call WriteRelatedProductTable(mclsProduct)
	If cblnAddon_ProductReview Then Call WriteProductReviewTable(mclsProduct)
	
	Call WriteCustomTable
	Call WriteProductSalesTable(mclsProduct)
	Call WriteSEOTable
	
	Call WriteFooterTable
	%>

</td>
</tr>
</table>
<%
End Sub	'WriteProductDetail

'************************************************************************************************************************************

Sub WriteGeneralTable

Dim i
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblGeneral">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label">Product ID:</td>
        <td><input id=prodID onchange='MakeDirty(this);' name=prodID Value="<%= mclsProduct.prodID %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label">Product Name:</td>
        <td><input id=prodName onchange="MakeDirty(this);" onblur="if (frmData.prodNamePlural.value == ''){frmData.prodNamePlural.value = frmData.prodName.value + 's'}" name=prodName Value="<%= EncodeString(mclsProduct.prodName,True) %>" maxlength=255 size=50></td>
      </tr>
      <tr>
        <td class="Label">Product Name (plural):</td>
        <td><input id=prodNamePlural onchange="MakeDirty(this);" name=prodNamePlural Value="<%= EncodeString(mclsProduct.prodNamePlural,True) %>" maxlength=255 size=50></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodShortDescription" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Short Description:</label></td>
        <td>
          <textarea id=prodShortDescription onchange="MakeDirty(this);" name=prodShortDescription rows="5" cols="50" title="Short Description" onkeyup="return checkMaxLength(this, 255, 'prodShortDescriptionCounter');"><%= mclsProduct.prodShortDescription %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.prodShortDescription);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a><span id="prodShortDescriptionCounter">(<%= Len(mclsProduct.prodShortDescription) %>/255)</span>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="prodDescription" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Long Description:</label></td>
        <td>
          <textarea id=prodDescription onchange="MakeDirty(this);" name=prodDescription rows="5" cols="50" title="Long Description"><%= mclsProduct.prodDescription %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.prodDescription);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a>
        </td>
      </tr>
      <tr>
        <td class="Label">Pricing:</td>
        <td>
            <input type=checkbox id=prodEnabledIsActive onchange="MakeDirty(this);" name=prodEnabledIsActive <% If mclsProduct.prodEnabledIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodEnabledIsActive">Is Active</label>
			<table border="1" cellpadding="2" cellspacing="0" class="tbl">
			  <tr class="tblhdr"><th>&nbsp;</th><th>Price</th><th <% If mclsProduct.prodSaleIsActive Then Response.Write " class=""Selected"" style=""color:black""" %>>Sale Price</th></tr>
			  <tr>
			    <td>&nbsp;</td>
			    <td><input id=prodPrice onchange="MakeDirty(this);" name=prodPrice Value="<%= mclsProduct.prodPrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
			    <td><input id=prodSalePrice onchange="MakeDirty(this);" name=prodSalePrice Value="<%= mclsProduct.prodSalePrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;<input type=checkbox id=prodSaleIsActive onchange="MakeDirty(this);" name=prodSaleIsActive <% If mclsProduct.prodSaleIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodSaleIsActive">Sale Is Active</label></td>
			  </tr>
			  <%
				For i = 0 To (mlngNumPricingLevels - 1)
					Response.Write "<tr>"
					Response.Write "<td>" & maryPricingLevels(i) & "</td>"

					maryPLPrices = Split(mclsProduct.prodPLPrice & "",";")
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = Trim(maryPLPrices(i))
					End If
					Response.Write "<td><INPUT id=prodPLPrice name=prodPLPrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"

					maryPLPrices = Split(mclsProduct.prodPLSalePrice & "",";")
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = Trim(maryPLPrices(i))
					End If
					Response.Write "<td><INPUT id=prodPLSalePrice name=prodPLSalePrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
					Response.Write "</tr>"

				Next 'i
			  %>
			</table>
        </td>
      </tr>
	  <% If cblnSF5AE Then %>
      <tr>
        <td class="Label">Gift Wrap Charge:</td>
        <td><input id=gwPrice onchange="MakeDirty(this);" name=gwPrice Value="<%= mclsProduct.gwPrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;
            <input type=checkbox name=gwActivate id=gwActivate onchange="MakeDirty(this);" <% WriteCheckboxValue mclsProduct.gwActivate %> value="ON">&nbsp;Gift Wrap Is Active
        </td>
      </tr>
	  <% End If %>
      <tr>
        <td class="Label"><label for="buyersClubPointValue" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Club Points:</label></td>
        <td><input name="buyersClubPointValue" id="buyersClubPointValue" onchange="MakeDirty(this);" Value="<%= mclsProduct.buyersClubPointValue %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;
            <input type=checkbox name="buyersClubIsPercentage" id="buyersClubIsPercentage" onchange="MakeDirty(this);" <% WriteCheckboxValue mclsProduct.buyersClubIsPercentage %> value="ON">&nbsp;Is Percentage?
        </td>
      </tr>

      <tr>
        <td class="Label"><label for="prodHandlingFee" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Handling Fee:</label></td>
        <td><input id="prodHandlingFee" onchange="MakeDirty(this);" name=prodHandlingFee Value="<%= mclsProduct.prodHandlingFee %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodSetupFee" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Setup Fee (Each item):</label></td>
        <td><input id="prodSetupFee" onchange="MakeDirty(this);" name=prodSetupFee Value="<%= mclsProduct.prodSetupFee %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodSetupFeeOneTime" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Setup Fee (Per item):</label></td>
        <td><input id="prodSetupFeeOneTime" onchange="MakeDirty(this);" name=prodSetupFeeOneTime Value="<%= mclsProduct.prodSetupFeeOneTime %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodMinQty" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Min Qty:</label></td>
        <td><input id="prodMinQty" onchange="MakeDirty(this);" name=prodMinQty Value="<%= mclsProduct.prodMinQty %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodIncrement" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Fraction:</label></td>
        <td><input id="prodIncrement" onchange="MakeDirty(this);" name=prodIncrement Value="<%= mclsProduct.prodIncrement %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'></td>
      </tr>

      <tr>
        <td class="Label">&nbsp;</td>
        <td>
        <input type=checkbox id=prodStateTaxIsActive onchange="MakeDirty(this);" name=prodStateTaxIsActive <% If mclsProduct.prodStateTaxIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label for="prodStateTaxIsActive">Apply State Tax to this item</label></td>
      </tr>
      <tr>
        <td class="Label">&nbsp;</td>
        <td>
        <input type=checkbox id=prodCountryTaxIsActive onchange="MakeDirty(this);" name=prodCountryTaxIsActive <% If mclsProduct.prodCountryTaxIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label for="prodCountryTaxIsActive">Apply Country Tax to this item</label></td>
      </tr>
      <tr>
        <td class="Label">Date Added:</td>
        <td><input name="prodDateAdded" id="prodDateAdded" onchange="MakeDirty(this);" value="<%= mclsProduct.prodDateAdded %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label">Date Modified:</td>
        <td><%= mclsProduct.prodDateModified %></td>
      </tr>
</table>
<%
End Sub	'WriteGeneralTable

'************************************************************************************************************************************

Function MakeImageHTTPSFriendly(strImage)

	If LCase(Request.ServerVariables("HTTPS")) = "on" Then
		If LCase(Left(strImage, 5)) = "http:" Then
			MakeImageHTTPSFriendly = Replace(strImage, "http:", "https:", 1, 1)
		Else
			MakeImageHTTPSFriendly = strImage
		End If
	Else
		MakeImageHTTPSFriendly = strImage
	End If

End Function	'MakeImageHTTPSFriendly

'************************************************************************************************************************************

Sub WriteDetailTable

Dim paryAdditionalImageText
Dim paryAdditionalImage
Dim paryAdditionalImageDesc
Dim plngImageCounter

%>
<span id=spantempFile style="display:none">
<input type=file id=tempFile name=tempFile onchange="ProcessPath(this);" size="20">
</span>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblDetail">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
		<% For plngImageCounter = 0 To UBound(maryImageFields) %>
		<tr>
			<% If mblnShortsImageManager Then %>
			<td class="Label"><span style="cursor:hand" onclick="PickImage('<%= maryImageFields(plngImageCounter)(1) %>');" title="Select image using Image Manager"><%= maryImageFields(plngImageCounter)(0) %></span>:</td>
			<% Else %>
			<td class="Label"><%= maryImageFields(plngImageCounter)(0) %>:</td>
			<% End If %>
			<td><input name="<%= maryImageFields(plngImageCounter)(1) %>" id="<%= maryImageFields(plngImageCounter)(1) %>" ondblclick="setDblClickDefault(this);MakeDirty(this);" onchange="MakeDirty(this);" Value="<%= maryImageFields(plngImageCounter)(2) %>" maxlength=255 size=60>
				<img style="cursor:hand" name="img<%= maryImageFields(plngImageCounter)(1) %>" id="img<%= maryImageFields(plngImageCounter)(1) %>" border="0" 
					onmouseover="DisplayTitle(this);showFullImage(this,0);return false;" onmouseout"showFullImage(this,1);return ClearTitle();" src="<%= MakeImageHTTPSFriendly(SetImagePath(maryImageFields(plngImageCounter)(2))) %>" 
					onclick="return SelectImage(this);" 
					title="Click to edit this image">
			</td>
		</tr>
		<% Next 'plngImageCounter %>
		<tr>
		  <td class="Label" valign="top">Additional Detail Images:</td>
		  <td>
		    <fieldset>
		      <legend>Additional Image Display Options</legend>
				<input type="radio" value="0" name="prodDisplayAdditionalImagesInWindow" id="prodDisplayAdditionalImagesInWindow0" <%= isChecked(mclsProduct.prodDisplayAdditionalImagesInWindow=0) %>><label for="prodDisplayAdditionalImagesInWindow0">Links to new window</label>&nbsp;
				<input type="radio" value="1" name="prodDisplayAdditionalImagesInWindow" id="prodDisplayAdditionalImagesInWindow1" <%= isChecked(mclsProduct.prodDisplayAdditionalImagesInWindow=1) %>><label for="prodDisplayAdditionalImagesInWindow1">Replace Detail Image</label>&nbsp;
				<input type="radio" value="2" name="prodDisplayAdditionalImagesInWindow" id="prodDisplayAdditionalImagesInWindow2" <%= isChecked(mclsProduct.prodDisplayAdditionalImagesInWindow=2) %>><label for="prodDisplayAdditionalImagesInWindow2">Show all</label>&nbsp;
		    </fieldset>
			<table class="tbl" width="100%" cellpadding="2" cellspacing="0" border="1" id="tblDetailImages">
			  <tr class="tblhdr">
			    <th>Image Title</th>
			    <th>Image Path</th>
			    <th>Image Description</th>
			  </tr>
			  <% For plngImageCounter = 0 To mclsProduct.NumAdditionalImages - 1 %>
			  <%
					paryAdditionalImageText = mclsProduct.AdditionalImageText
					paryAdditionalImage = mclsProduct.AdditionalImage
					paryAdditionalImageDesc = mclsProduct.AdditionalImageDesc
			  %>
			  <tr>
			    <td><input type=text name=additionalImageText<%= plngImageCounter %> ID=additionalImageText<%= plngImageCounter %> value="<%= paryAdditionalImageText(plngImageCounter) %>" size=40></td>
			    <td><input type=text name=additionalImage<%= plngImageCounter %> ID=additionalImage<%= plngImageCounter %> value="<%= paryAdditionalImage(plngImageCounter) %>" size=40></td>
			    <td><input type=text name=additionalImageDesc<%= plngImageCounter %> ID=additionalImageDesc<%= plngImageCounter %> value="<%= paryAdditionalImageDesc(plngImageCounter) %>" size=40></td>
			  </tr>
			  <% Next 'plngImageCounter %>
			  <tr>
			    <td><input type=text name=additionalImageText<%= plngImageCounter %> ID=additionalImageText<%= plngImageCounter %> value="" size=40></td>
			    <td><input type=text name=additionalImage<%= plngImageCounter %> ID=additionalImage<%= plngImageCounter %> value="" size=40></td>
			    <td><input type=text name=additionalImageDesc<%= plngImageCounter %> ID=additionalImageDesc<%= plngImageCounter %> value="" size=40></td>
			  </tr>
			</table>
		  </td>
		</tr>
      <tr>
        <td class="Label"><span style="cursor:hand"title="Automatically set link to detail.asp">Link:</span></td>
        <td><input name="prodLink" id="prodLink" ondblclick="setDblClickDefault(this);MakeDirty(this);"  onchange="MakeDirty(this);" Value="<%= EncodeString(mclsProduct.prodLink,True) %>" maxlength=255 size=60>&nbsp;<img src="images/preview.gif" title="view this page" style="cursor:hand" onclick="OpenHelp('../../../' + document.frmData.prodLink.value)"></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodMessage" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Confirmation Message:</label></td>
        <td>
          <textarea id=prodMessage onchange="MakeDirty(this);" name=prodMessage rows="5" cols="50" title="Confirmation Message"><%= mclsProduct.prodMessage %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.prodMessage);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a>
        </td>
      </tr>
      <tr>
        <td class="Label">Category</td>
        <td>
			<select size="1"  id=prodCategoryId name=prodCategoryId onchange="MakeDirty(this);">
			<% 'Call MakeCombo(mrsCategory,"catName","catID",mclsProduct.prodCategoryId) %>
			<%= MakeCombo_Saved("category", mclsProduct.prodCategoryId) %>
			</select>
		</td>        
      </tr>
      <tr>
        <td class="Label"><a href="sfManufacturersAdmin.asp">Manufacturer</a>:</td>
        <td>
			<select size="1"  id=prodManufacturerId name=prodManufacturerId onchange="MakeDirty(this);">
			<% 'Call MakeCombo(mrsManufacturer,"mfgName","mfgID",mclsProduct.prodManufacturerId) %>
			<%= MakeCombo_Saved("manufacturer", mclsProduct.prodManufacturerId) %>
			</select>&nbsp;<input type="text" name="ManufacturerNew" id="ManufacturerNew" value="" maxlength="50" title="Enter the name of the manufacturer you wish to create">
		</td>        
      </tr>
      <tr>
        <td class="Label"><a href="sfVendorsAdmin.asp">Vendor</a>:</td>
        <td>
			<select size="1"  id=prodVendorId name=prodVendorId onchange="MakeDirty(this);">
			<% 'Call MakeCombo(mrsVendor,"vendName","vendID",mclsProduct.prodVendorId) %>
			<%= MakeCombo_Saved("vendor", mclsProduct.prodVendorId) %>
			</select>&nbsp;<input type="text" name="VendorNew" id="VendorNew" value="" maxlength="50" title="Enter the name of the vendor you wish to create">
		</td>        
      </tr>
      </table>
      </td>
      </tr>
</table>
<%
End Sub	'WriteDetailTable 

'************************************************************************************************************************************

Sub WriteShippingTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblShipping">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label"><label for="prodShip" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Product Based Shipping Cost:</label></td>
        <td><input id=prodShip onchange="MakeDirty(this);" name=prodShip Value="<%= mclsProduct.prodShip %>" size=6>&nbsp;
            <input type=checkbox id=prodShipIsActive onchange="MakeDirty(this);" name=prodShipIsActive <% If mclsProduct.prodShipIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodShipIsActive">This item is shipped</label>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="prodFixedShippingCharge" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Fixed Shipping Cost:</label></td>
        <td><input name="prodFixedShippingCharge" id="prodFixedShippingCharge" onchange="MakeDirty(this);" Value="<%= mclsProduct.prodFixedShippingCharge %>" size=6></td>
      </tr>
      <tr>
        <td class="Label"><label for="prodSpecialShippingMethods" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Enabled Special Shipping Methods:</label></td>
        <td><input name="prodSpecialShippingMethods" id="prodSpecialShippingMethods" onchange="MakeDirty(this);" Value="<%= mclsProduct.prodSpecialShippingMethods %>" size=20> <a href="ssPostageRate_shippingMethodsAdmin.asp" target="_blank">Click here to see shipping methods</a></td>
      </tr>
      <tr>
        <td class="Label">Weight:</td>
        <td><input id=prodWeight onchange="MakeDirty(this);" name=prodWeight Value="<%= mclsProduct.prodWeight %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Height:</td>
        <td><input id=prodHeight onchange="MakeDirty(this);" name=prodHeight Value="<%= mclsProduct.prodHeight %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Width:</td>
        <td><input id=prodWidth onchange="MakeDirty(this);" name=prodWidth Value="<%= mclsProduct.prodWidth %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Length:</td>
        <td><input id=prodLength onchange="MakeDirty(this);" name=prodLength Value="<%= mclsProduct.prodLength %>" size=6></td>
      </tr>
      <tr>
        <td colspan="2"><hr /></td>
      </tr>
      <tr>
        <th colspan="2">Electronic Download</th>
      </tr>
      <tr>
        <td class="Label"><label for="prodFileName" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">File Path:</label></td>
        <td><input id=prodFileName onchange="MakeDirty(this);" name=prodFileName value="<%= mclsProduct.prodFileName %>" maxlength=255 size=60>
		</td>
      </tr>
      <tr>
        <td class="Label"><label for="prodMaxDownloads" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Max. Number of Downloads:</label></td>
        <td><input id="prodMaxDownloads" onchange="MakeDirty(this);" name=prodMaxDownloads value="<%= mclsProduct.prodMaxDownloads %>" maxlength=10 size=10>
		</td>
      </tr>
      <tr>
        <td class="Label"><label for="prodDownloadValidFor" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Days Download Valid For:</label></td>
        <td><input id="prodDownloadValidFor" onchange="MakeDirty(this);" name=prodDownloadValidFor value="<%= mclsProduct.prodDownloadValidFor %>" maxlength=10 size=10>
		</td>
      </tr>

      <tr>
        <td class="Label"><label for="version">Version:</label></td>
        <td><%= writeHTMLFormElement(enDisplayType_textbox, "25", "version", "version", mclsProduct.version, "", " onchange=""MakeDirty(this);""") %></td>
      </tr>
      <tr>
        <td class="Label"><label for="releaseDate">Release Date:</label></td>
        <td><%= writeHTMLFormElement(enDisplayType_textbox_WithDateSelect, "10", "releaseDate", "releaseDate", mclsProduct.releaseDate, "", " onchange=""MakeDirty(this);""") %></td>
      </tr>
       <tr>
        <td class="Label"><label for="InstallationHours">Installation Hours:</label></td>
        <td>
        <%= writeHTMLFormElement(enDisplayType_textbox, "10", "InstallationHours", "InstallationHours", mclsProduct.InstallationHours, "", " onchange=""MakeDirty(this);""") %>
        <%= writeHTMLFormElement(enDisplayType_checkbox, "10", "InstallationRequired", "InstallationRequired", mclsProduct.InstallationRequired, "", " onchange=""MakeDirty(this);""") %> <label for="InstallationRequired">Installation Required</label>
        </td>
      </tr>
      <tr>
        <td class="Label"></td>
        <td><%= writeHTMLFormElement(enDisplayType_checkbox, "10", "MyProduct", "MyProduct", mclsProduct.MyProduct, "", " onchange=""MakeDirty(this);""") %> <label for="MyProduct">My Product</label></td>
      </tr>
      <tr>
        <td class="Label"></td>
        <td><%= writeHTMLFormElement(enDisplayType_checkbox, "10", "IncludeInSearch", "IncludeInSearch", mclsProduct.IncludeInSearch, "", " onchange=""MakeDirty(this);""") %> <label for="IncludeInSearch">Include In Search Results</label></td>
      </tr>
      <tr>
        <td class="Label"></td>
        <td><%= writeHTMLFormElement(enDisplayType_checkbox, "10", "IncludeInRandomProduct", "IncludeInRandomProduct", mclsProduct.IncludeInRandomProduct, "", " onchange=""MakeDirty(this);""") %> <label for="IncludeInRandomProduct">Include In Random Product</label></td>
      </tr>
      <tr>
        <td class="Label"><label for="UpgradeVersion">Upgrade Version:</label></td>
        <td><%= writeHTMLFormElement(enDisplayType_textbox, "10", "UpgradeVersion", "UpgradeVersion", mclsProduct.UpgradeVersion, "", " onchange=""MakeDirty(this);""") %></td>
      </tr>
      <tr>
        <td class="Label"><label for="packageCodes">Package Codes:</label></td>
        <td><%= writeHTMLFormElement(enDisplayType_textbox, "10", "packageCodes", "packageCodes", mclsProduct.packageCodes, "", " onchange=""MakeDirty(this);""") %></td>
      </tr>
</table>
<%
End Sub	'WriteShippingTable 

'************************************************************************************************************************************

Sub WriteAttributeTable

Dim i
Dim plngAttrSize
Dim plngAttrdtSize

	If isObject(mclsProduct.rsAttributes) Then
		mclsProduct.rsAttributes.Filter = "attrProdId='" & mclsProduct.prodID & "'"
		plngAttrSize = mclsProduct.rsAttributes.recordcount + 1
	Else
		plngAttrSize = 3
	End If

%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblAttributes">
      <tr>
        <td>&nbsp;</td>
        <td>
        <table class="tbl" width=100% border=0 ID="Table5">
        <tr>
        <td valign=top>
			<table class="tbl" border=0 cellpadding=0 cellspacing=0 ID="Table6">
			  <tr>
				<td valign=top>
					<select size="<%= plngAttrSize %>" id=attrID name=attrID onchange="ChangeAttr(this);">
					<option value="">Create New Attribute Category</option>
					<%
						If isObject(mclsProduct.rsAttributes) Then 
							mclsProduct.rsAttributes.Filter = "attrProdId='" & mclsProduct.prodID & "'"
							Call MakeCombo(mclsProduct.rsAttributes,"attrName","attrID",mclsProduct.attrID)
						End If
					%>
					</select>
				 </td>
				 <td valign=middle>
<% If mclsProduct.AttributeCategoryOrderable Then %>
					<br>
					<input type=image id=imgUp1 src="images/up.gif" onclick="UpItem('attribute'); return false;" title="Move Attribute Up" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" NAME="imgUp1">
					<br><input type=image id=imgDown1 src="images/down.gif" onclick="DownItem('attribute'); return false;" title="Move Attribute Down" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" NAME="imgDown1">
<% Else %>
&nbsp;
<% End If %>
				 </td>
			</tr>
			</table>

		</td>
		<td>&nbsp;</td>

		<td>
			<table class="tbl">
			  <tr>
			    <td class="label">Attribute Category:&nbsp;</td>
			    <td><input name="attrName" id="attrName" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrName)  %>" maxlength=50 size=50></td>
			  </tr>
			<% If mclsProduct.TextBasedAttribute Then %>
			  <% If Len(mclsProduct.attrDisplay_Field) > 0 Then  %>
			  <tr>
			    <td class="label"><span title="The contents of this field are written out exactly as it appears here instead of the attribute cateory name.">Custom Display Text:&nbsp;</span></td>
			    <td>
			    <input name="attrDisplay" id="attrDisplay" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrDisplay)  %>" maxlength=255 size=50 onblur="this.size=50;" onfocus="this.size=100">
				<a href="javascript:doNothing()" onclick="return openACE(document.frmData.attrDisplay);" title="Edit this field with the HTML Editor">
				<img src="images/prop.bmp" border=0></a>
			    </td>
			  </tr>
			  <% End If %>
			  <% If Len(mclsProduct.attrSKU_Field) > 0 Then  %>
			  <tr>
			    <td class="label"><span title="Attribute category specific portion of the SKU.">SKU:&nbsp;</span></td>
			    <td><input name="attrSKU" id="attrSKU" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrSKU)  %>" size=20 maxlength=50></td>
			  </tr>
			  <% End If %>
			<tr>
			<td class="label" valign="top">Display Style:&nbsp;</td>
			<td>
				<% If True Then %>
				<select name="attrDisplayStyle" id="attrDisplayStyle">
					<% For i = 0 To UBound(maryAttributeTypes) %>
					<option value="<%= i %>" <%= isChecked((mclsProduct.attrDisplayStyle = i) OR (Len(mclsProduct.attrDisplayStyle & "") = 0)) %>><%= maryAttributeTypes(i) %></option>
					<% Next 'i %>
				</select>
				<% Else %>
				<table class="tbl" border=1 cellspacing=0 cellpadding=2 ID="Table7">
					<% For i = 0 To UBound(maryAttributeTypes) %>
					<tr>
					<td align=left><input type="radio" value="<%= i %>" <% If (mclsProduct.attrDisplayStyle = i) OR (Len(mclsProduct.attrDisplayStyle & "") = 0) Then Response.Write "checked" %> id="attrDisplayStyle<%= i %>" name="attrDisplayStyle" onchange='MakeDirty(this);'>&nbsp;<label for="attrDisplayStyle<%= i %>"><%= maryAttributeTypes(i) %></label></td>
					</tr>
					<% Next 'i %>
				</table>
				<% End If %>
			</td>
			</tr>
			<% End If 'pblnTextBasedAttribute %>
		<% If Len(mclsProduct.attrImage_Field) > 0 And True Then %>
		<tr>
			<% If mblnShortsImageManager Then %>
			<td class="Label"><span style="cursor:hand" onclick="PickImage('attrImage');" title="Select attribute image using Image Manager">Attribute Image</SPAN>:</td>
			<% Else %>
			<td class="Label"><span title="This field is used to specify the image path for the attribute onchange selection.">Image Path:</span></td>
			<% End If %>
			<td><input name="attrImage" id="attrImage" onchange="MakeDirty(this);" value="<%= mclsProduct.attrImage %>" maxlength=255 size=60>
				<img style="cursor:hand" name="imgattrImage" id="Img1" border="0" 
					onmouseover="DisplayTitle(this);" onmouseout"return ClearTitle();" src="<%= SetImagePath(mclsProduct.attrImage) %>" 
					onclick="return SelectImage(this);" 
					title="Click to edit the Product image">
			</td>
		</tr>
		<% End If 'Len(mclsProduct.attrdtImage_Field) > 0 %>
			  <% If Len(mclsProduct.attrURL_Field) > 0 Then  %>
			  <tr>
			    <td class="label"><span title="This field changes the attribute category name to a hyperlink. The contents of this field are added immediately after the href= attribute. It is NOT surrounded by quotes.">URL:&nbsp;</span></td>
			    <td><input name="attrURL" id="attrURL" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrURL)  %>" maxlength=255 size=50 onblur="this.size=50;" onfocus="this.size=100"></td>
			  </tr>
			  <% End If %>
			  <% If Len(mclsProduct.attrExtra_Field) > 0 Then  %>
			  <tr>
			    <td class="label"><span title="Unused.">Extra:&nbsp;</span></td>
			    <td><input name="attrExtra" id="attrExtra" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrExtra)  %>" size=50 maxlength=50></td>
			  </tr>
			  <% End If %>
			</table>

		
		</td>
		<td>&nbsp;</td>
		<td align=center>
			<input class='butn' title="Delete this attribute category" id=btnDeleteAttr name=btnDeleteAttr type=button value='Delete Category' onclick='var theForm = this.form; var blnConfirm=confirm("Are you sure you wish to delete attribute category " + document.frmData.attrName.value + "?"); if (blnConfirm){theForm.Action.value = "DeleteAttribute"; theForm.submit();}'><br>
			<input class='butn' title="Delete all attributes" id="btnDeleteAllAttr" name=btnDeleteAllAttr type=button value='Delete Attributes' onclick='var theForm = this.form; var blnConfirm=confirm("Are you sure you wish to delete all attributes from this product?"); if (blnConfirm){theForm.Action.value = "DeleteAllAttributes"; theForm.submit();}'><br>
			<input class='butn' title="Copy this attribute category and attributes to a new category for this product" id=btnDuplicateAttr name=btnDuplicateAttr type=button value='Duplicate Category' onclick="var theForm = this.form; var pstrNewprodName = prompt('Enter New Attribute Category Name','New Attribute Category');if (pstrNewprodName != null){theForm.CopyProduct.value = pstrNewprodName;theForm.Action.value = 'DuplicateAttr';if (<%= LCase(CStr(cblnSF5AE)) %>){theForm.ChangeInventory.value = true;}theForm.submit();}" disabled><br>
			<input class='butn' title="Copy just this attribute category and attributes to an existing product" id=btnCopyAttr name=btnCopyAttr type=button value='Copy Category' onclick='var theForm = this.form; var pstrNewprodName = prompt("Enter Product ID to copy attribute to","Enter Product ID");if (pstrNewprodName != null){theForm.CopyProduct.value = pstrNewprodName;theForm.Action.value = "CopyAttr";theForm.submit();}' disabled>
			<input class='butn' title="Copy all of this product's attribute category and attributes to an existing product" id=btnCopyProduct name=btnCopyProduct type=button value='Copy Attributes' onclick='var theForm = this.form; var pstrNewprodName = prompt("Enter Product ID to copy attributes to","Enter Product ID"); if (pstrNewprodName != null){ theForm.CopyProduct.value = pstrNewprodName; theForm.Action.value = "CopyAttributesToProd"; theForm.submit();}'>&nbsp;
		</td>
		</tr>
		</table>
		</td>        
      </tr>
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanAttrOptions"><%= mclsProduct.attrName %> &nbsp;</span></th>
  </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
        <table class="tbl" width=100% border=0>
        <tr>
        <td valign=top>
			<table class="tbl" border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td valign=middle>
				  <div id="divAttrOptions">&nbsp;</div>
				  <select size=3  id=attrdtID name=attrdtID onchange="ChangeAttrDetail(this);">
				  <option value="">Create New Attribute</option>
				  </select>
				 </td>
				 <td valign=middle>
					<br>
					<input type=image id=imgUp src="images/up.gif" onclick="UpItem('attributeDetail'); return false;" title="Move Attribute Up" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" disabled NAME="imgUp">
					<br><input type=image id=imgDown src="images/down.gif" onclick="DownItem('attributeDetail'); return false;" title="Move Attribute Down" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" disabled NAME="imgDown">
				 </td>
			</tr>
			</table>
		</td>
		<td valign=top>&nbsp;</td>
		<td>
		<table class="tbl">
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td class="Label">Attribute:</td>
			<td><input id=attrdtName onchange='if (<%= LCase(CStr(cblnSF5AE)) %>){this.form.ChangeInventory.value = true;MakeDirty(this);}' name=attrdtName value="<%= mclsProduct.attrdtName  %>" maxlength=50 size=50></td>
		</tr>
		<% If Len(mclsProduct.attrdtDisplay_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="The contents of this field are written out exactly as it appears here instead of the attribute detail name.">Custom Display Text:</span></td>
			<td>
			<input name="attrdtDisplay" id="attrdtDisplay" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtDisplay %>" maxlength=255 size=50 onblur="this.size=50;" onfocus="this.size=100">
			<a href="javascript:doNothing()" onclick="return openACE(document.frmData.attrdtDisplay);" title="Edit this field with the HTML Editor">
			<img src="images/prop.bmp" border=0></a>
			</td>
		</tr>
		<% End If 'Len(pstrattrdtDisplay_Field) > 0 %>
		<% If Len(mclsProduct.attrdtSKU_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="SKU">SKU:</span></td>
			<td><input name="attrdtSKU" id="attrdtSKU" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtSKU %>" maxlength=25 size=25></td>
		</tr>
		<% End If 'Len(attrdtSKU_Field) > 0 %>
		<tr>
			<td class="Label">Price Variance:</td>
			<td>
			<input id=attrdtPrice onchange='MakeDirty(this);if (this.value==0){this.form.attrdtType[2].checked=true;}else{this.form.attrdtType[0].checked=true;}' name=attrdtPrice value="<%= mclsProduct.attrdtPrice  %>" size="6">&nbsp;
			<% If False Then %>
			<select name="attrdtType" id="attrdtType">
				<option value="1" <%= isChecked(mclsProduct.attrdtType=1) %>>Increase</option>
				<option value="2" <%= isChecked(mclsProduct.attrdtType=2) %>>Decrease</option>
				<option value="0" <%= isChecked(mclsProduct.attrdtType=0) %>>No Change</option>
			</select>
			<% Else %>
			<input type="radio" name="attrdtType" id="attrdtType1" value="1" <%= isChecked(mclsProduct.attrdtType=1) %>><label for="attrdtType1">Increase</label>
			<input type="radio" name="attrdtType" id="attrdtType2" value="2" <%= isChecked(mclsProduct.attrdtType=2) %>><label for="attrdtType2">Decrease</label>
			<input type="radio" name="attrdtType" id="attrdtType0" value="0" <%= isChecked(mclsProduct.attrdtType=0) %>><label for="attrdtType0">No Change</label>
			<% End If %>
			</td>
		</tr>
		<% If mblnAttrPrice Then %>
		<tr>
		<td class="Label">&nbsp;</td>
		<td>
			<table class="tbl" border=1 cellspacing=0 cellpadding=0 ID="Table14">
				<%= mstrHeaderRow %>
				<tr>
				<%
				maryPLPrices = Split(mclsProduct.attrdtPLPrice & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = Trim(maryPLPrices(i))
					End If
					Response.Write "<td align=center><INPUT id=attrdtPLPrice name=attrdtPLPrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
				%>
				</tr>
			</table>
		</td>
		</tr>
		<% End If	'mblnAttrPrice %>
		<% If False Then %>
		<tr>
			<td>&nbsp;</td>
			<td>
			<input type="radio" value="1" <% if mclsProduct.attrdtType=1 then Response.Write "Checked" %> id="attrdtType1" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType1">Increase</label><br>
			<input type="radio" value="2" <% if mclsProduct.attrdtType=2 then Response.Write "Checked" %> id="attrdtType2" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType2">Decrease</label><br>
			<input type="radio" value="0" <% if mclsProduct.attrdtType=0 then Response.Write "Checked" %> id="attrdtType0" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType0">No Change</label>
			</td>
		</tr>
		<% End If %>
		<% If Len(mclsProduct.attrdtWeight_Field) > 0 And True Then %>
		<tr>
			<td class="Label">Weight Variance:</td>
			<td><input name="attrdtWeight" id="attrdtWeight" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtWeight %>" maxlength=255 size=10></td>
		</tr>
		<% End If 'Len(pstrattrdtWeight_Field) > 0 %>
		<% If Len(mclsProduct.attrdtImage_Field) > 0 And True Then %>
		<tr>
			<% If mblnShortsImageManager Then %>
			<td class="Label"><span style="cursor:hand" onclick="PickImage('attrdtImage');" title="Select attribute image using Image Manager">Attribute Image</SPAN>:</td>
			<% Else %>
			<td class="Label"><span title="This field is used to specify the image path for the attribute onchange selection.">Image Path:</span></td>
			<% End If %>
			<td><input name="attrdtImage" id="attrdtImage" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtImage %>" maxlength=255 size=60>
				<img style="cursor:hand" name="imgattrdtImage" id="imgattrdtImage" border="0" 
					onmouseover="DisplayTitle(this);" onmouseout"return ClearTitle();" src="<%= SetImagePath(mclsProduct.attrdtImage) %>" 
					onclick="return SelectImage(this);" 
					title="Click to edit the Product image">
			</td>
		</tr>
		<% End If 'Len(mclsProduct.attrdtImage_Field) > 0 %>
		<% If Len(mclsProduct.attrdtURL_Field) > 0 Then  %>
		<tr>
		<td class="label"><span title="This field changes the attribute detail name to a hyperlink. The contents of this field are added immediately after the href= attribute. It is NOT surrounded by quotes.">URL:&nbsp;</span></td>
		<td><input name="attrdtURL" id="attrdtURL" onchange="MakeDirty(this);" value="<%= Server.HTMLEncode(mclsProduct.attrdtURL)  %>" maxlength=255 size=50 onblur="this.size=50;" onfocus="this.size=100"></td>
		</tr>
		<% End If	'Len(mclsProduct.attrdtURL_Field) > 0 %>
		<% If Len(mclsProduct.attrdtFileName_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="File Name">File Name:</span></td>
			<td><input name="attrdtFileName" id="attrdtFileName" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtFileName %>" maxlength=255 size=50></td>
		</tr>
		<% End If 'Len(pstrattrdtFileName_Field) > 0 %>
		<% If Len(mclsProduct.attrdtExtra_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="Extra use">Extra:</span></td>
			<td><input name="attrdtExtra" id="attrdtExtra" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtExtra %>" maxlength=255 size=50></td>
		</tr>
		<% End If 'Len(pstrattrdtExtra_Field) > 0 %>
		<% If Len(mclsProduct.attrdtExtra1_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="Extra use">Extra:</span></td>
			<td><input name="attrdtExtra1" id="attrdtExtra1" onchange="MakeDirty(this);" value="<%= mclsProduct.attrdtExtra1 %>" maxlength=255 size=50></td>
		</tr>
		<% End If 'Len(pstrattrdtExtra1_Field) > 0 %>
		<% If Len(mclsProduct.attrdtDefault_Field) > 0 And True Then %>
		<tr>
			<td class="Label"><span title="Check this to automatically select this attribute.">Select by default:</span></td>
			<td><input type=checkbox name="attrdtDefault" id="attrdtDefault" onchange="MakeDirty(this);" value="1" <%= isChecked(mclsProduct.attrdtDefault) %>></td>
		</tr>
		<% End If 'Len(attrdtDefault_Field) > 0 %>
		</table>
		</td>
		<td>&nbsp;</td>
		<td align=center>
			<input class='butn' id=btnDeleteAttrDetail name=btnDeleteAttrDetail type=button value='Delete Attribute' onclick='var theForm = this.form; var blnConfirm = confirm("Are you sure you wish to delete attribute " + document.frmData.attrdtName.value + "?"); if (blnConfirm){theForm.Action.value = "DeleteAttrDetail"; theForm.submit();}' disabled><br>
		</td>
		</tr>
		</table>
		</td>        
      </tr>
</table>
<%
End Sub	'WriteAttributeTable

'************************************************************************************************************************************

 Sub WriteInventoryTable 

 Dim i
 Dim plngRecordCount
 Dim prsInventory
 Dim paryInventory
 
	On Error Resume Next
	With mclsProduct.rsInventoryInfo
%>
<input type=hidden id=ChangeInventory name=ChangeInventory value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblInventory">
      <tr>
        <td>&nbsp;</td>
        <td>
        <table class="tbl" width=100% border=0 ID="Table11">
          <tr>
            <td class="Label"><label id='lblinvenInStockDEF' name='lblinvenInStockDEF' for='invenInStockDEF' class="Label">Default Inventory Qty</label>:</td>
            <td>
              <input type='text' id='invenInStockDEF' name='invenInStockDEF' value='<%= .Fields("invenInStockDEF").value %>' onblur='return isInteger(this, false, "Please enter an integer for the quantity");' size="20">
              <input type='checkbox' id='invenbTracked' name='invenbTracked' <% WriteCheckboxValue .Fields("invenbTracked").value %> value="ON">&nbsp;<label id='lblinvenbTracked' name='lblinvenbTracked' for='invenbTracked'>Track Inventory</label>
            </td>
          </tr>
          <tr>
            <td class="Label"><label id='lblinvenLowFlagDEF' name='lblinvenLowFlagDEF' for='lblinvenLowFlagDEF'>Default Notify Qty</label>:</td>
            <td>
              <input type='text' id='invenLowFlagDEF' name='invenLowFlagDEF' value='<%= .Fields("invenLowFlagDEF").value %>' onblur='return isInteger(this, false, "Please enter an integer for the quantity");' size="20">
              <input type='checkbox' id='invenbNotify' name='invenbNotify' <% WriteCheckboxValue .Fields("invenbNotify").value %> value="ON">&nbsp;<label id='lblinvenbNotify' name='lblinvenbNotify' for='invenbNotify'>Notify when stock reaches this level</label>
            </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>
              <input type='checkbox' id='invenbBackOrder' name='invenbBackOrder' <% WriteCheckboxValue .Fields("invenbBackOrder").value %> value="ON">&nbsp;<label id='lblinvenbBackOrder' name='lblinvenbBackOrder' for='invenbBackOrder'>Allow Back Order</label><br>
			  <input type='checkbox' id='invenbStatus' name='invenbStatus' <% WriteCheckboxValue .Fields("invenbStatus").value %> value="ON">&nbsp;<label id='lblinvenbStatus' name='lblinvenbStatus' for='invenbStatus'>Show Stock Status on Search Page</label>
            </td>
          </tr>
		</table>
		</td></tr>
		<tr><td colspan=5><hr></td></tr>
		<tr>
		  <td colspan=5 align="center">
		  <table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblInventoryLevels">
<%
	End With	'mclsProduct.rsInventoryInfo

'Call writeTimer("Getting inventory details")
	Set prsInventory = GetRS("Select invenId, invenAttName, invenInStock, invenLowFlag from sfInventory Where invenProdId='" & SQLSafe(mclsProduct.prodID) & "'")
'Call writeTimer("Inventory details opened")
	If prsInventory.EOF Then
		plngRecordCount = -1
	Else
		plngRecordCount = prsInventory.RecordCount -1
	End If
	
	If plngRecordCount > -1 Then
	%>
			<tr>
			  <th colspan=3 align=left>Inventory Records: <%= plngRecordCount + 1 %></th>
			</tr>
			<tr>
			  <th>Attribute</th>
			  <th>Qty In Stock</th>
			  <th>Notify When Qty Reaches</th>
			</tr>
	<%
Response.Flush
		paryInventory = prsInventory.GetRows()
		'invenId, invenAttName, invenInStock, invenLowFlag
	End If
	Call ReleaseObject(prsInventory)
'Call writeTimer("GetRows Complete")

	For i = 0 To plngRecordCount
		Response.Write "<tr><input type='hidden' id='invenId' name='invenId' value='" & paryInventory(0, i) & "'>" _
					 & "<td class=Label>" & paryInventory(1, i) & ":&nbsp;&nbsp;</td>" _
					 & "<td align=center><input type=text name=invenInStock id=invenInStock value='" & paryInventory(2, i) & "' onblur=""return isInteger(this, true, 'Please enter an integer for the quantity');"" onchange=""ChangeAE('Inventory');"" size=20></td>" _
					 & "<td align=center><input type=text name=invenLowFlag id=invenLowFlag value='" & paryInventory(3, i) & "' onblur=""return isInteger(this, true, 'Please enter an integer for the quantity');"" onchange=""ChangeAE('Inventory');"" size=20></td>" _
					 & "</tr>"
	Next 'i

%>
		  </table>
		  </td></tr>
		</TD>        
      </TR>
</table>
<%
End Sub	'WriteInventoryTable

'************************************************************************************************************************************

Sub WriteMTPTable

Dim i
%>
<script>
function AddMTP()
{
var pNewRow;
var pNewCell;
var cstrQuote = '"';
var pstrCell1 = "<input type='text' id='mtQuantity' name='mtQuantity' value='0' onblur='return isInteger(this, true, " + cstrQuote + "Please enter an integer for the quantity" + cstrQuote + ");'>"
var pstrCell2 = "<input type='text' id='mtValue' name='mtValue' value='0' onblur='return isNumeric(this, true, " + cstrQuote + "Please enter an integer for the discount" + cstrQuote + ");'>"
var pstrCell3 = "<select id='mtType' name='mtType'><option>Amount</option><option>Percent</option></select>"      
var pstrCell4 = "<INPUT class='butn' id=btnDeleteMTP name=btnDeleteMTP type='button' value='Delete Discount Level' onclick='DeleteMTP();'>" 

	pNewRow = document.all("tblMTPInput").insertRow();
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell1;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell2;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell3;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell4;
	
}

function DeleteMTP(theCell)
{
var ptheRow = window.event.srcElement.parentElement.parentElement;
ptheRow.parentElement.deleteRow(ptheRow.rowIndex);
}

</script>
<input type=hidden id=ChangeMTP name=ChangeMTP value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblMTP">
<tr>
<td><input type=checkbox name="prodLimitQtyToMTP" id="prodLimitQtyToMTP" onchange="MakeDirty(this);" <%= isChecked(mclsProduct.prodLimitQtyToMTP) %> value="ON">&nbsp;<label FOR="prodLimitQtyToMTP">Limit quantity selections to pricing breaks</label>
</td>
</tr>
<tr><td>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblMTPInput">
      <tr>
        <td>Quantity</td>
        <td>Discount Amount</td>
        <td>Discount Type<br><i>(checked: value, unchecked: percentage)</i></td>
        <td>&nbsp;</td>
      </tr>
<% 
'mtProdID
'mtIndex
'mtQuantity
'mtValue
'mtType

Dim pblnAmount
Dim j

If isObject(mclsProduct.rsMTP) Then
  With mclsProduct.rsMTP
	For j = 1 To .RecordCount
%>
      <tr>
        <td valign=top><input type='text' id='mtQuantity' name='mtQuantity' value='<%= .Fields("mtQuantity").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");' onchange="ChangeAE('MTP');"></td>
		<td valign=top>
		  <input type='text' id='mtValue' name='mtValue' value='<%= .Fields("mtValue").value %>' onblur='return isNumeric(this, true, "Please enter an integer for the discount");' onchange="ChangeAE('MTP');">
			<% If mblnMTPrice Then %>
			<table class="tbl" border=1 cellspacing=0 cellpadding=0 ID="Table15">
			 <%= mstrHeaderRow %>
             <tr>
             <%
				maryPLPrices = Split(.Fields("mtPLValue").value & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = Trim(maryPLPrices(i))
					End If
					Response.Write "<td align=center><INPUT id=mtPLValue" & j & " name=mtPLValue" & j & " onchange=" & Chr(34) & "ChangeAE('MTP');" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
             %>
             </tr>
			</table>
			<% End If %>
		</td>
		<% If (.Fields("mtType").value="Amount") Then %>
		<td valign=top><select id='mtType' name='mtType' onchange="ChangeAE('MTP');"><option selected>Amount</option><option>Percent</option></select></td>        
		<% Else %>
		<td valign=top><select id="Select1" name='mtType' onchange="ChangeAE('MTP');"><option>Amount</option><option selected>Percent</option></select></td>        
		<% End If %>
		<td valign=top><input class='butn' id=btnDeleteMTP name=btnDeleteMTP type='button' value='Delete Discount Level' onclick="DeleteMTP(this); ChangeAE('MTP');"></td>        
      </tr>
<%	
	mclsProduct.rsMTP.MoveNext 
	Next 'j
  End With
End If
%>
</table>
</td><tr>
<tr>
	<td><input class='butn' id=btnNewMTP name=btnNewMTP type='button' value='New Discount Level' onclick="AddMTP(); ChangeAE('MTP');"></td></tr>
</table>
<% 
End Sub	'WriteMTPTable 

'************************************************************************************************************************************

Sub WriteCategoryTable
%>

<input type=hidden id=ChangeCategory name=ChangeCategory value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblCategory">
  <tr>
    <td>
      <table class="tbl" width="100%" border="0" ID="Table12">
		<tr>
		  <th>Categories</th>
		  <th>&nbsp;</th>
		  <th>This product is in the following categories:</th>
		</tr>
		<tr>
		  <td align=center><%= mclsProduct.CategoryList %></td>
		  <td valign=middle align=center>
			<input class="butn" type=button id="btnAddCategory" name="btnAddCategory" onclick="AddCategory(); ChangeAE('Category');" value="-->"><br>
			<input class="butn" type=button id="btnDeleteCategory" name="btnDeleteCategory" onclick="DeleteCategory(); ChangeAE('Category');" value="<--"><br>
			<input class='butn' title="Copy categories to an existing product" id="btnCopyCategories" name="btnCopyCategories" type=button value='Copy Categories' onclick='var theForm = this.form; var pstrNewprodName = prompt("Enter Product ID to copy categories to","Enter Product ID");if (pstrNewprodName != null){theForm.CopyProduct.value = pstrNewprodName;theForm.Action.value = "CopyCategories";theForm.submit();}'>
		  </td>
		  <td align=center>
			<select id=Categories name=Categories size=10 multiple>
			</select>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% 
End Sub	'WriteCategoryTable

'************************************************************************************************************************************

Sub WriteCustomTable

Dim i
Dim paryCustomValues

paryCustomValues = mclsProduct.CustomValues
If Not isArray(paryCustomValues) Then Exit Sub
%>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblCustom">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
	<% For i = 0 To UBound(paryCustomValues) %>
      <tr>
        <td class="Label"><%= paryCustomValues(i)(0) %>:</td>
        <td>
        <%= writeHTMLFormElement(paryCustomValues(i)(3), paryCustomValues(i)(4), paryCustomValues(i)(1), paryCustomValues(i)(1), paryCustomValues(i)(2), paryCustomValues(i)(5), "MakeDirty(this);") %>
        </td>
      </tr>
    <% Next 'i %>
</table>
<% 
End Sub	'WriteCustomTable 

'************************************************************************************************************************************

Sub WriteRelatedProductTable(byRef objclsProduct)
%>
<script language="javascript">
var mdicrelatedProducts = new ActiveXObject("Scripting.Dictionary");
<%
	Response.Write setCustomDictionary(objclsProduct.relatedProducts, ";", "relatedProducts", "product")
%>
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblRelatedProducts">
  <tr>
	<td>
	  <select id="relatedProducts" name="relatedProducts" size="5" ondblclick="openMovementWindow('relatedProducts','product');" multiple></select>
	  <a href="" onclick="openMovementWindow('relatedProducts','product'); return false;"><img src="images/properites.gif" border="0"></a>
	</td>
  </tr>
</table>
<% 
End Sub	'WriteRelatedProductTable 

'************************************************************************************************************************************

Sub WriteProductReviewTable(byRef objclsProduct)
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblProductReview">
    <tr>
    <td align=left>
    <input type=checkbox name=prodEnableReviews id=prodEnableReviews onchange="MakeDirty(this);" <%= isChecked(mclsProduct.prodEnableReviews) %> value="ON">&nbsp;<label for="prodEnableReviews">Enable Reviews for this item</label></td>
  </tr>
  <tr>
	<td>
<%
Dim i
Dim paryProductReviews
	'0-contentReferenceID
	'1-contentContent
	'2-contentAuthorName
	'3-contentAuthorEmail
	'4-contentAuthorShowEmail
	'5-contentAuthorRating
	'6-contentDateCreated
Dim pstrAuthor
Dim pstrURL

	If loadProductReviews(objclsProduct.prodID, paryProductReviews) Then
%>
	<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" ID="Table17">
	  <tr>
		<td class="tdContent2">
		<%
		If mlngNumReviews = 0 Then
			Response.Write "This product has not been rated yet."
		Else
			If mlngNumReviews = 1 Then
				Response.Write "One reviewer gave this a rating of " & ratingDisplayImage(mlngAvgReviewScore)
			Else
				Response.Write mlngNumReviews & " reviewers gave this an average rating of " & ratingDisplayImage(mlngAvgReviewScore)
			End If

			For i = 0 To mlngNumReviews - 1
			
				If Abs(paryProductReviews(4, i)) = 1 Then
					pstrAuthor = "<a href='mailTo:" & Trim(paryProductReviews(3, i)) & "'>" & Trim(paryProductReviews(2, i)) & "</a>"
				Else
					pstrAuthor = Trim(paryProductReviews(2, i))
				End If
				
				Response.Write "<hr />"
				Response.Write pstrAuthor & " " & ratingDisplayImage(paryProductReviews(5, i)) & "<br />"
				If Len(Trim(paryProductReviews(1, i))) > 0 Then
					Response.Write "<strong>Comments:</strong> " & Trim(paryProductReviews(1, i)) & "<br />"
				End If
				Response.Write "<strong>Date Reviewed:</strong> " & FormatDateTime(paryProductReviews(6, i)) & "<br />"
				Response.Write "<a href=ssProductReviewsAdmin.asp?Action=viewItem&viewID=" & paryProductReviews(0, i) & ">Edit</a><br />"
				
				If paryProductReviews(7, i) > 0 Or paryProductReviews(8, i) > 0 Then
					Response.Write paryProductReviews(7, i) & " of " & paryProductReviews(7, i) + paryProductReviews(8, i) & " found this review useful.<br />"
				End If
			Next 'i
		End If
		%>
		</td>
	  </tr>
	</table>
<%
	Else
		Response.Write "This product has not been rated yet."
	End If	'Not loadProductReviews
%>
	</td>
  </tr>
</table>
<% 
End Sub	'WriteProductReviewTable 

'************************************************************************************************************************************

Sub WriteProductSalesTable(byRef objclsProduct)

Dim pstrSQL
Dim pobjRS

	pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderAmount, sfOrderDetails.odrdtQuantity" _
			& " FROM sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
			& " WHERE sfOrderDetails.odrdtProductID=" & wrapSQLValue(objclsProduct.prodID, False, enDatatype_string) _
			& " ORDER BY orderID DESC"
	Set pobjRS = GetRS(pstrSQL)
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblProductSales">
<colgroup align="center"></colgroup>
<colgroup align="center"></colgroup>
<colgroup align="center"></colgroup>
<colgroup align="center"></colgroup>
    <tr>
    <td align=left colspan=4>
    <input type=checkbox name=prodEnableAlsoBought id=prodEnableAlsoBought onchange="MakeDirty(this);" <%= isChecked(mclsProduct.prodEnableAlsoBought) %> value="ON">&nbsp;<label for="prodEnableAlsoBought">Enable customers who bought this also bought display for this item</label></td>
  </tr>
  <tr class="tblhdr">
	<th>Order #</th>
	<th>Date</th>
	<th>Order Amount</th>
	<th>Qty of this item</th>
  </tr>
  <% Do While Not pobjRS.EOF %>
  <tr>
	<td><a href="ssOrderAdmin.asp?Action=ViewOrder&OrderID=<%= pobjRS.Fields("orderID").Value %>"><%= pobjRS.Fields("orderID").Value %></a></td>
	<td><%= customFormatDateTime(pobjRS.Fields("orderDate").Value, 1, "-") %></td>
	<td><%= WriteCurrency(pobjRS.Fields("orderAmount").Value) %></td>
	<td><%= pobjRS.Fields("odrdtQuantity").Value %></td>
  </tr>
  <%
		pobjRS.MoveNext
	 Loop
	 Call ReleaseObject(pobjRS)
  %>
</table>
<% 
End Sub	'WriteProductSalesTable 

'**************************************************************************************************************************************************

Sub WriteFooterTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
		<input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/ProductManager/help_ProductManager.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
        <input class='butn' title='Create a new product' id=btnNewProduct name=btnNewProduct type=button value='New' onclick='return btnNewProduct_onclick(this)'>&nbsp;
        <input class='butn' title='Create a new product based on this product' id=btnDuplicateProduct name=btnDuplicateProduct type=button value='Duplicate' onclick='DuplicateProduct(this);'>&nbsp;
        <input class='butn' title="Reset" id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)' disabled>&nbsp;&nbsp;
        <input class='butn' title="Delete this product" id=btnDeleteProduct name=btnDeleteProduct type=button value='Delete' onclick='return btnDeleteProduct_onclick(this)'>
        <input class='butn' title="Save changes" id=btnUpdate name=btnUpdate type=button value='Save Changes' onclick='return ValidInput(this.form);'>
    </td>
  </tr>
</table>
<%
End Sub	'WriteFooterTable

'**************************************************************************************************************************************************

Function WriteCheckboxValue(vntValue)

	If len(Trim(vntValue) & "") > 0 Then
		If cBool(vntValue) Then Response.Write "CHECKED"
	End If


End Function	'WriteCheckboxValue

'************************************************************************************************************************************

Sub WriteCustomMTPTable
%>
<script>
function AddPricingLevel()
{
var pNewRow;
var pNewCell;
var cstrQuote = '"';
var pstrCell1 = "<input type='text' id='PricingLevel' name='PricingLevel' value='0' onblur='return isInteger(this, true, " + cstrQuote + "Please enter an integer for the quantity" + cstrQuote + ");'>"
var pstrCell2 = "<input type='text' id='PricingAmount' name='PricingAmount' value='0' onblur='return isNumeric(this, true, " + cstrQuote + "Please enter an integer for the discount" + cstrQuote + ");'>"
var pstrCell3 = "<INPUT class='butn' id=btnDeletePricingLevel name=btnDeletePricingLevel type='button' value='Delete Pricing Level' onclick='DeletePricingLevel();'>" 

	pNewRow = document.all("tblPricingLevelInput").insertRow();
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell1;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell2;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell3;
	
}

function DeletePricingLevel(theCell)
{
var ptheRow = window.event.srcElement.parentElement.parentElement;
ptheRow.parentElement.deleteRow(ptheRow.rowIndex);
}

</script>
<input type=hidden id=ChangePricingLevel name=ChangePricingLevel value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="Table16">
<tr><td>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblPricingLevelInput">
      <tr>
        <td>Quantity</td>
        <td>Price</td>
        <td>&nbsp;</td>
      </tr>
<% 
Dim pblnAmount

If isObject(mclsProduct.rsPricingLevels) Then
  With mclsProduct.rsPricingLevels
	Do While Not .EOF
%>
      <tr>
        <td><input type='text' id='PricingLevel' name='PricingLevel' value='<%= .Fields("PricingLevel").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");'"></td>
		<td><input type='text' id='PricingAmount' name='PricingAmount' value='<%= .Fields("PricingAmount").value %>' onblur='return isNumeric(this, true, "Please enter an integer for the discount");'"></td>
		<td><input class='butn' id=btnDeletePricingLevel name=btnDeletePricingLevel type='button' value='Delete Pricing Level' onclick="DeletePricingLevel(this);"></td>        
      </tr>
<%	
	mclsProduct.rsPricingLevels.MoveNext 
	Loop
  End With
End If
%>
</table>

</td><tr>
<tr>
	<td><input class='butn' id=btnNewPricingLevel name=btnNewPricingLevel type='button' value='New Pricing Level' onclick="AddPricingLevel();"></td></tr>
</table>
<%

End Sub	'WriteCustomMTPTable

'************************************************************************************************************************************

Sub WriteSEOTable

%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblSEO">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label"><label for="pageName" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Page Name:</label></td>
        <td>
          <textarea name="pageName" id="pageName" onchange="MakeDirty(this);" rows="2" cols="50" title="Page Name" onkeyup="return checkMaxLength(this, 100, 'pageNameCounter');"><%= mclsProduct.pageName %></textarea>
          <span id="pageNameCounter">(<%= Len(mclsProduct.pageName) %>/100)</span>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="metaTitle" onmouseover="showDataEntryTip(this);" onmouseout="htm();">meta Title:</label></td>
        <td>
          <textarea name="metaTitle" id="metaTitle" onchange="MakeDirty(this);" rows="2" cols="50" title="meta Title" onkeyup="return checkMaxLength(this, 100, 'metaTitleCounter');"><%= mclsProduct.metaTitle %></textarea>
          <span id="metaTitleCounter">(<%= Len(mclsProduct.metaTitle) %>/100)</span>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="metaDescription" onmouseover="showDataEntryTip(this);" onmouseout="htm();">meta Description:</label></td>
        <td>
          <textarea name="metaDescription" id="metaDescription" onchange="MakeDirty(this);" rows="6" cols="50" title="meta Description" onkeyup="return checkMaxLength(this, 255, 'metaDescriptionCounter');"><%= mclsProduct.metaDescription %></textarea>
          <span id="metaDescriptionCounter">(<%= Len(mclsProduct.metaDescription) %>/255)</span>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="metaKeywords" onmouseover="showDataEntryTip(this);" onmouseout="htm();">meta Keywords:</label></td>
        <td>
          <textarea name="metaKeywords" id="metaKeywords" onchange="MakeDirty(this);" rows="2" cols="50" title="meta Keywords" onkeyup="return checkMaxLength(this, 100, 'metaKeywordsCounter');"><%= mclsProduct.metaKeywords %></textarea>
          <span id="metaKeywordsCounter">(<%= Len(mclsProduct.metaKeywords) %>/100)</span>
        </td>
      </tr>
</table>
<% End Sub	'WriteSEOTable %>
