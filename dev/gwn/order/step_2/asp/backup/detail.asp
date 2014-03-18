<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeout = 900

	'@DESCRIPTION: Product detail Page

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/ssProductReview.asp"-->
<!--#include file="include_files/detail_imageDisplay.asp"-->
<%

'Remove comment if you only want logged in customers to be able to view
'If Not isLoggedIn Then Response.Redirect "myAccount.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME") & Server.URLEncode("?" & Request.QueryString)
On Error Goto 0
'Variable declarations
Dim txtProdId, SQL
Dim attrName, attrNamePrev, icounter, strOut, iAttrNum, strAttrPrice
Dim mstrProductDescription

	txtProdId = Request.QueryString("product_id")
	
	Call setRecentlyViewedProducts(txtProdId, Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString)

	If getProductInfo(txtProdId, enProduct_Exists) Then
		If getProductInfo(txtProdId, enProduct_AttrNum) > 0 Then
			SQL = "SELECT sfAttributes.*, sfAttributeDetail.* " _
				& "FROM sfAttributes INNER JOIN sfAttributeDetail ON sfAttributes.attrId = sfAttributeDetail.attrdtAttributeId " _
				& "WHERE attrProdId = '" & makeInputSafe(txtProdId) & "'" _
				& "ORDER BY attrDisplayOrder, AttrName, attrdtOrder"
			'Response.Write "SQL: " & SQL & "<br />"

		End If
	End If	'getProductInfo(txtProdId, enProduct_Exists)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= getProductInfo(txtProdId, enProduct_metaTitle) %></title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="keywords" content="<%= getProductInfo(txtProdId, enProduct_metaKeywords) %>">
<meta name="description" content="<%= getProductInfo(txtProdId, enProduct_metaDescription) %>">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">

<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/ssAttributeExtender.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
function validateForm(theForm)
{
	if (theForm.QUANTITY.type == "text"){theForm.QUANTITY.quantityBox=true;}
	if (theForm.QUANTITY.type == "select-one"){theForm.selQUANTITY.optional=true;}

	return sfCheck(theForm);
}

<% If getProductInfo(txtProdId, enProduct_Exists) Then Response.Write "prodBasePrice =" & getProductInfo(txtProdId, enProduct_SellPrice) & ";" & vbcrlf %>
</script>
<% writeCurrencyConverterOpeningScript %>
</head>
<body <%= mstrBodyStyle %> onload="theCustomImage.src = getCustomImagePath();">

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="tdbackgrnd">
        <tr>
          <td class="tdContent2">
          <%
    If getProductInfo(txtProdId, enProduct_Exists) Then
		Call DisplayProductEditLink(txtProdId)
		mstrProductDescription = getProductInfo(txtProdId, enProduct_Description)
		If Len(mstrProductDescription) = 0 Then mstrProductDescription = getProductInfo(txtProdId, enProduct_Description)
		%>
		<!--<div class="clsCategoryTrail" id="divCategoryTrail"><% Call writeDetailCategoryTrail(txtProdId, True) %></div>-->
		<!--webbot bot="PurpleText" PREVIEW="Begin Optional Confirmation Message Display" -->
		<% Call WriteThankYouMessage %>
		<!--webbot bot="PurpleText" PREVIEW="End Optional Confirmation Message Display" -->
		<form method="post" name="<%= MakeFormNameSafe(txtProdId) %>" action="<%= C_HomePath %>addproduct.asp" onSubmit="return validateForm(this);">
		<input TYPE="hidden" NAME="PRODUCT_ID" VALUE="<%= txtProdId %>">         
		<table border="0" width="100%" class="tdContent2" cellpadding="2" cellspacing="0">
		  <tr>
			<td align="center" valign="top"><%= detailImageOut %></td>
			<td align="left" valign="top">
				<h1 class="productName"><%= getProductInfo(txtProdId, enProduct_Name) %></h1>
				<% If False Then %>
				<strong><%= C_ProductID %>:</strong>&nbsp;<%= txtProdId %><br />
				<% End If %>
				<% If getProductInfo(txtProdId, enProduct_MfgID) <> 1 Then %>
					<strong><%= C_ManufacturerNameS %>:</strong>&nbsp;<a href="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "URL", True) %>" title="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "MetaTitle", True) %>"><%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "Name", True) %></a><br />
				<% End If %>
				<% If getProductInfo(txtProdId, enProduct_VendorID) <> 1 Then %>
					<strong><%= C_VendorNameS %>:</strong>&nbsp;<a href="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "URL", False) %>" title="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "MetaTitle", False) %>"><%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "Name", False) %></a><br />
				<% End If %>
				<strong><%= C_Description %>:</strong>&nbsp;<%= mstrProductDescription %><br />
				 
				<%
				If hasMTP(txtProdId) Then
					Response.Write "<div align=""center"">"
					Response.Write WriteMTPricingTable(txtProdId, "Price Per Pair")
					Response.Write "</div>"
				Else
				    Response.Write "<strong>" & C_Price & ":</strong>&nbsp;"
				    If iConverion = 1 Then
					    If getProductInfo(txtProdId, enProduct_SaleIsActive) Then 
						    Response.Write "<span class=""itemOnSalePrice""><script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_Price)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
						    Response.Write "<span class=""SalesPrice"">" & C_SPrice & ": <script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_SalePrice)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_SalePrice)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
						    Response.Write "<span class=""YouSave"">" & C_YSave & ": <script>document.write(""" & FormatCurrency(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & " = ("" + OANDAconvert(" & trim(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
					    Else
						    Response.Write "<script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_Price)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"
					    End If 
				    Else
					    If getProductInfo(txtProdId, enProduct_SaleIsActive) Then 
						    Response.Write "<span class=""itemOnSalePrice"">" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & "</span><br />"
						    Response.Write "<span class=""SalesPrice"">" & C_SPrice & ": " & FormatCurrency(getProductInfo(txtProdId, enProduct_SalePrice)) & "</span><br />"
						    Response.Write "<span class=""YouSave"">" & C_YSave & ": " & FormatCurrency(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & "</span><br />"
					    Else
						    Response.Write FormatCurrency(getProductInfo(txtProdId, enProduct_Price))
					    End If
				    End If	'iConverion = 1
				End If

				'If cblnSF5AE Then
				'	SearchResults_GetProductInventory txtProdId
				'	SearchResults_ShowMTPricesLink txtProdId
				'End If
				%>
				<!--<br /><a href="Shipping Calculator" onclick="var mScreenHeight = window.screen.availHeight; var mScreenWidth = window.screen.availWidth; window.open('ssShippingEstimator.asp?ProductID=<%= Server.URLEncode(txtProdId) %>','SearchResults','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=' + mScreenHeight/1.5 + ',width=' + mScreenWidth/2 + ',screenY=' + mScreenHeight/1.5 + ',screenX=' + mScreenWidth/2 + ',top=0,left=' + mScreenWidth/2.05 + ',resizable'); return false;">Estimate Shipping</a>-->
				<%
				Call DebugRecordTime("Display attributes")
				' -----------------------------
				' ATTRIBUTE OUTPUT ::: BEGIN --
				' -----------------------------
				If getProductInfo(txtProdId, enProduct_AttrNum) > 0 Then
					Response.Write "<br /><table border=0 cellspacing=0 cellpadding=0>"
					Call DisplayAttributes_New(strOut, MakeFormNameSafe(txtProdId))
					Response.Write strOut
					Response.Write "</table>"
				End If 
				
				' ------------------------
				' ATTRIBUTE OUTPUT ::: END
				' ------------------------
				%>
				<% If getProductInfo(txtProdId, enProduct_IsActive) Then %>
				 
				<% If getProductInfo(txtProdId, enProduct_LimitQtyToMTP) Then %>
				<table class="quantityDisplay">
				<tr><th colspan="3" class="quantityDisplay">Select a Quantity</th></tr>
				<% 
				'<input type="hidden" name="QUANTITY" id="QUANTITY" value="1">
				'<select name="selQUANTITY" id="selQUANTITY" onchange="this.form.QUANTITY.value=this.options[this.options.selectedIndex].value;">
				'</select>
					Dim maryMTP
					Dim i

					maryMTP = getProductInfo(txtProdId, enProduct_MTP)
					For i = 0 To UBound(maryMTP)
						'Response.Write "<option value=""" & maryMTP(i)(0) & """>" & maryMTP(i)(0) & "</option>"
						Response.Write "<tr>"
						Response.Write "<td width=""15px""><input type=""radio"" name=""QUANTITY"" id=""QUANTITY" & i & """ value=""" & maryMTP(i)(0) & """></td>"
						Response.Write "<td width=""30px"" align=""right""><label for=""QUANTITY" & i & """>" & maryMTP(i)(0) & "</label>&nbsp;</td>"
						Response.Write "<td>&nbsp;<label for=""QUANTITY" & i & """>Your Price: " & CustomCurrency(calculateMTPrice(i)) & "</label></td>"
						Response.Write "</tr>"
					Next 'i
				%>
				</table>
				<% ElseIf Not mblnQtyBoxAttributeUsed Then %>
				<p>Quantity:<input class="formDesign" type="text" name="QUANTITY" title="Quantity" size="3" value="1" onfocus="this.select();" onchange="return validateQty(this, '<%= getProductInfo(txtProdId, enProduct_prodMinQty) %>');"></p>
				<% End If %>
				<%
				  Dim prodIncrement: prodIncrement = getProductInfo(txtProdId, enProduct_prodIncrement)
				  If prodIncrement > 0 Then
					If prodIncrement > 1 Then
						Response.Write "&nbsp;<select name=""QUANTITY_FRACTION." & txtProdId & """ id=""QUANTITY_FRACTION"">"
						Response.Write "<option value=""0"">0 / " & prodIncrement & "</option>"
						For icounter = 1 To prodIncrement - 1
							Response.Write "<option value=""" & icounter/prodIncrement & """>" & icounter & " / " & prodIncrement & "</option>"
						Next 'icounter
						Response.Write "</select>"
					End If
				  End If
				%>

				<% If cblnSF5AE Then SearchResults_GetGiftWrap txtProdId 'SFAE%>
				<p><input type="image" class="inputImage" name="AddProduct" src="<%= C_BTN03 %>" alt="Add To Cart">
				<% End If	'getProductInfo(txtProdId, enProduct_IsActive) %>
				<% If iSaveCartActive = 1 Then %><br /><input type="image" class="inputImage" name="SaveCart" src="<%= C_BTN02 %>" alt="Save To Cart"><% End If %>
				<% If iEmailActive = 1 Then %><br /><a href="javascript:emailFriend('<%= txtProdId %>')"><img border="0" src="<%= C_BTN24 %>" alt="Email a Friend"></a><% End If %>
				</p>
		      </td>
		    </tr>
		  </table>
		</form>
		<%= WriteJavaScript(mstrssAttributeExtenderjsOut) %>

<!--webbot bot="PurpleText" PREVIEW="Start Dynamic Product - People Who Bought This Also Bought" -->
<%
	Call DebugRecordTime("Load also bought")
	If getProductInfo(txtProdId, enProduct_EnableAlsoBought) Then
	Set mclsDynamicProducts = New clsDynamicProducts
	With mclsDynamicProducts
		.DisplayType = 2			'0 - General Catalog; 1 - Best Sellers; 2 - Also Bought (requires current product ID); 3 - Related Products
		.SortField = "Price"	'Sort on this field to order by sale volume
		.TemplateName = "relatedProducts_Detail.htm"
		.CellStyle = "class=""contentCenter"""
		.WrapTable = True	'suppress table wrap since using a list style output
		.ImageNotPresentURL = "images/NoImage.gif"
		.Connection = cnn

		.CurrentProductID = txtProdId		'Only used for Also Bought displays
		.NumColumns = 3
		.NumRows = 1

		.CustomDateField = ""	'Must be blank to use to limit orders placed by a date range
		'.CustomDateAfter = DateAdd("m", -12, Date())
		'Example, orders placed only in the last 12 months

		If .LoadDynamicProducts Then
		%>
		<table border="1" cellpadding="0" cellspacing="0" width="100%" class="Section">
		<tr>
			<th class="tdMiddleTopBanner">People who bought this product also purchased</th>
		</tr>
		<tr>
			<td><!--webbot bot="PurpleText" PREVIEW="Start Dynamic Product" --><% .DisplayDynamicProducts %><!--webbot bot="PurpleText" PREVIEW="End Dynamic Product" --></td>
		</tr>
		</table>
		<%		
		End If	'LoadDynamicProducts
	End With	'mclsDynamicProducts
	End If	'getProductInfo(txtProdId, enProduct_EnableAlsoBought)
%>
<!--webbot bot="PurpleText" PREVIEW="End Dynamic Product - People Who Bought This Also Bought" -->

<!--webbot bot="PurpleText" PREVIEW="Start Dynamic Product - Related Products" -->
<%
	Call DebugRecordTime("Load related products")
	'No reason to load this section if no related product
	If Len(getProductInfo(txtProdId, enProduct_RelatedProducts)) > 0 Then
	Set mclsDynamicProducts = New clsDynamicProducts
	With mclsDynamicProducts
		.DisplayType = 3			'0 - General Catalog; 1 - Best Sellers; 2 - Also Bought (requires current product ID); 3 - Related Products
		.SortField = "Name"	'Sort on this field to order by sale volume
		.TemplateName = "relatedProducts_Detail.htm"
		.CellStyle = "class=""contentCenter"""
		.WrapTable = True	'suppress table wrap since using a list style output
		.ImageNotPresentURL = "images/NoImage.gif"
		.DebugEnabled = False
		.ProductIDList = getProductInfo(txtProdId, enProduct_RelatedProducts)

		.DisplayType = 3			'0 - General Catalog; 1 - Best Sellers; 2 - Also Bought (requires current product ID); 3 - Related Products
		.RelatedProductField = "relatedProducts"	'Only used for Related Products displays
		.CurrentProductID = txtProdId		'Only used for Also Bought displays
		.NumColumns = 3
		.NumRows = 1

		.CustomDateField = ""	'Must be blank to use to limit orders placed by a date range
		'.CustomDateAfter = DateAdd("m", -12, Date())
		'Example, orders placed only in the last 12 months

		On Error Resume Next
		If isObject(cnn) Then .Connection = cnn
		If Err.number > 0 Then Err.Clear
		On Error Goto 0
		
		If .LoadDynamicProducts Then
		%>
		<br />
		<table border="1" cellpadding="0" cellspacing="0" width="100%" class="Section">
		<tr>
			<th class="tdMiddleTopBanner">You may also wish to consider</th>
		</tr>
		<tr>
			<td><!--webbot bot="PurpleText" PREVIEW="Start Dynamic Product" --><% .DisplayDynamicProducts %><!--webbot bot="PurpleText" PREVIEW="End Dynamic Product" --></td>
		</tr>
		</table>
		<%		
		End If	'LoadDynamicProducts
	End With	'mclsDynamicProducts
	End If	'Len(getProductInfo(txtProdId, enProduct_RelatedProducts)) > 0
	Set mclsDynamicProducts = Nothing
	%>
	<!--webbot bot="PurpleText" PREVIEW="End Dynamic Product - Related Products" -->
	<% If getProductInfo(txtProdId, enProduct_EnableReviews) Then %>
	<br />
	<center><% ssProductReview(txtProdId) %></center>
	<% End If	'getProductInfo(txtProdId, enProduct_EnableReviews) %>
      </td>
    </tr>
<% Else %>
        <tr>
          <td class="tdContent2">
			<table border=0 width="100%">
			  <tr><td align="center">Product <strong><%= txtProdId %></strong> was not found in the current product inventory</td></tr>
			  <tr><td width="100%" colspan="2"><hr noshade color="#000000" size="1" width="90%"></td></tr>
			</table>
			<!--#include file="include_files/detail_searchBox_notfound.asp"-->
		  </td>
		</tr>
<% End If	'getProductInfo(txtProdId, enProduct_Exists) %>
    </table>
</td>
</tr>
</table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
</body>
</html>
<%
	Call cleanup_dbconnopen
%>