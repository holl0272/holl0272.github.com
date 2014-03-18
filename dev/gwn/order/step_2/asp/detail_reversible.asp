<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeout = 900
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
Dim i

Dim mstrPrefix
Dim maryJerseyAttributes
Dim jerseyAttributeCounter
Dim numEntries
Dim EntryCounter
Dim maryAttributes
Dim attributeDetail
Dim attributeDetail2

Const cstrCustomLogo_ProductID = "11600"	'CustomLogo, 10400Logo

Const enJerseyAttribute_Name = 0
Const enJerseyAttribute_Position = 1
Const enJerseyAttribute_Default = 2
Const enJerseyAttribute_Type = 3	'0-hidden; 1-select; 2-hidden text; 3-text
Const cblnDebugJersey = False    'True   False

Dim enAttrPos_Size
Dim enAttrPos_Number
Dim enAttrPos_NameOnJersey
Dim enAttrPos_JerseyColor
Dim enAttrPos_LetteringOption
Dim enAttrPos_LetteringColor
Dim enAttrPos_LetteringFont
Dim enAttrPos_LetteringStyleTeam
Dim enAttrPos_LetteringStyleName
Dim enAttrPos_LocationTeam
Dim enAttrPos_TeamName
Dim enAttrPos_LetteringColor_Side2

	numEntries = request.form("jerseyRows")
	txtProdId = Request.QueryString("product_id")
	If getProductInfo(txtProdId, enProduct_Exists) Then
		Call setRecentlyViewedProducts(txtProdId, Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString)
		mstrPrefix = MPOPrefix(txtProdId)
		maryAttributes = getProductInfo(txtProdId, enProduct_attributes)
		Call LoadJerseyAttributes("Basketball")
	End If	'getProductInfo(txtProdId, enProduct_Exists)

'**********************************************************************************************************

Sub LoadJerseyAttributes(byVal JerseyType)

Dim attrName
Dim i, j
Dim paryAttributeDetails

	Select Case JerseyType
		Case "Basketball"
		    i = 0
			enAttrPos_JerseyColor = i:  i = i + 1
			enAttrPos_LetteringColor = i:  i = i + 1
			enAttrPos_LetteringColor_Side2 = i:  i = i + 1
			enAttrPos_LetteringOption = i:  i = i + 1
			enAttrPos_LocationTeam = i:  i = i + 1
			enAttrPos_TeamName = i:  i = i + 1
			enAttrPos_LetteringStyleTeam = i:  i = i + 1
			enAttrPos_LetteringFont = i:  i = i + 1
			enAttrPos_Size = i:  i = i + 1
			enAttrPos_NameOnJersey = i:  i = i + 1
			enAttrPos_Number = i:  i = i + 1
			enAttrPos_LetteringStyleName = i:  i = i + 1
			numEntries = request.form("jerseyRows")

			ReDim maryJerseyAttributes(i-1)
			'Attribute Name, Position, Default Value
			maryJerseyAttributes(enAttrPos_JerseyColor) =           Array("Jersey Color", -1, "", 0)
			maryJerseyAttributes(enAttrPos_LetteringColor) =        Array("Lettering Color", -1, "", 0)
			maryJerseyAttributes(enAttrPos_LetteringColor_Side2) =  Array("Lettering Color 2nd Side", -1, "", 0)
			maryJerseyAttributes(enAttrPos_LetteringOption) =       Array("Lettering Option", -1, "", 0)
			maryJerseyAttributes(enAttrPos_LocationTeam) =          Array("Location of Team Name", -1, "", 0)
			maryJerseyAttributes(enAttrPos_TeamName) =              Array("Team Name", -1, "", 2)
			maryJerseyAttributes(enAttrPos_LetteringStyleTeam) =    Array("Team Name Lettering Style", -1, "", 0)
			maryJerseyAttributes(enAttrPos_LetteringFont) =         Array("Lettering Font", -1, "", 0)
			maryJerseyAttributes(enAttrPos_Size) =                  Array("Size", -1, "", 1)
			maryJerseyAttributes(enAttrPos_NameOnJersey) =          Array("Player Name", -1, "", 3)
			maryJerseyAttributes(enAttrPos_Number) =                Array("Player Number", -1, "", 3)
			maryJerseyAttributes(enAttrPos_LetteringStyleName) =    Array("Player Name Lettering Style", -1, "", 0)

		Case "Football"
			ReDim maryJerseyAttributes(10)
			'Attribute Name, Position, Default Value
			maryJerseyAttributes(0) = Array("Jersey Color", -1, "", 0)
			maryJerseyAttributes(1) = Array("Lettering Color", -1, "", 0)
			maryJerseyAttributes(2) = Array("Lettering Option", -1, "", 0)
			maryJerseyAttributes(3) = Array("Location of Team Name", -1, "", 0)
			maryJerseyAttributes(4) = Array("Team Name", -1, "", 2)
			maryJerseyAttributes(5) = Array("Team Name Lettering Style", -1, "", 0)
			maryJerseyAttributes(6) = Array("Lettering Font", -1, "", 0)

			maryJerseyAttributes(7) = Array("Size", -1, "", 1)
			maryJerseyAttributes(8) = Array("Player Name", -1, "", 3)
			maryJerseyAttributes(9) = Array("Player Number", -1, "", 3)
			maryJerseyAttributes(10) = Array("Player Name Lettering Style", -1, "", 0)

			enAttrPos_JerseyColor = 0
			enAttrPos_LetteringColor = 1
			enAttrPos_LetteringOption = 2
			enAttrPos_LocationTeam = 3
			enAttrPos_TeamName = 4
			enAttrPos_LetteringStyleTeam = 5
			enAttrPos_LetteringFont = 6
			enAttrPos_Size = 7
			enAttrPos_NameOnJersey = 8
			enAttrPos_Number = 9
			enAttrPos_LetteringStyleName = 10
			numEntries = request.form("jerseyRows")

		Case Else
	End Select

	For jerseyAttributeCounter = 0 To UBound(maryJerseyAttributes)
		For i = 0 To UBound(maryAttributes)
			attrName = maryAttributes(i)(enAttribute_Name)
			If attrName = maryJerseyAttributes(jerseyAttributeCounter)(0) Then
				maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) = i + 1	'correct for being a 0 based array
				paryAttributeDetails = maryAttributes(i)(enAttribute_DetailArray)
				'Set the default value to the first entry; this covers the situation if no default has been set
				maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) = paryAttributeDetails(0)(enAttributeDetail_ID)
				For j = 0 To UBound(paryAttributeDetails)
					If paryAttributeDetails(j)(enAttributeDetail_Default) Then
						maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) = paryAttributeDetails(j)(enAttributeDetail_ID)
						Exit For
					End If
				Next 'j
				Exit For
			End If
		Next 'i
	Next 'jerseyAttributeCounter

	If cblnDebugJersey Then
		Response.Write "<fieldset style=""color:black;background:white""><legend>Jersey Attributes</legend>"
		For jerseyAttributeCounter = 0 To UBound(maryJerseyAttributes)
			response.Write "Attribute (" & jerseyAttributeCounter & "): " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Name) & " - " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & " - " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & "<br />"
		Next 'jerseyAttributeCounter
		Response.Write "<hr>"
		Response.Write "Jersey Color Position: " & enAttrPos_JerseyColor & "<br />"
		Response.Write "Size Position: " & enAttrPos_Size & "<br />"
		Response.Write "Number Position: " & enAttrPos_Number & "<br />"
		Response.Write "Name On Jersey Position: " & enAttrPos_NameOnJersey & "<br />"
		Response.Write "Lettering Style Name: " & enAttrPos_LetteringStyleName & "<br />"
		Response.Write "</fieldset>"
	End If

End Sub	'LoadJerseyAttributes

'**********************************************************************************************************

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
<script language="javascript" src="SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

function validateForm(theForm)
{
	if (theForm.QUANTITY.type == "text"){theForm.QUANTITY.quantityBox=true;}
	if (theForm.QUANTITY.type == "select-one"){theForm.selQUANTITY.optional=true;}

	return sfCheck(theForm);
}

<% If getProductInfo(txtProdId, enProduct_Exists) Then Response.Write "prodBasePrice =" & getProductInfo(txtProdId, enProduct_SellPrice) & ";" & vbcrlf %>

(function ( $ ) {

    $.fn.jsonTable = function( options ) {
        var settings = $.extend({
            head: [],
            json:[]
        }, options, { table: this } );

        table = this;

        table.data("settings",settings);

        if (table.find("thead").length == 0) {
            table.append($("<thead></thead>").append("<tr></tr>"));
        }

        if (table.find("thead").find("tr").length == 0) {
            table.find("thead").append("<tr></tr>");
        }

        if (table.find("tbody").length == 0) {
            table.append($("<tbody></tbody>"));
        }

        $.each(settings.head, function(i, header) {
            table.find("thead").find("tr").append("<th>"+header+"</th>");
        });

        return table;
    };

    $.fn.jsonTableUpdate = function( options ){
        var opt = $.extend({
            source: undefined,
            rowClass: undefined,
            callback: undefined
        }, options );
        var settings = this.data("settings");

        if(typeof opt.source == "string")
        {
            $.get(opt.source, function(data) {
                $.fn.updateFromObj(data,settings,opt.rowClass, opt.callback);
            });
        }
        else if(typeof opt.source == "object")
        {
            $.fn.updateFromObj(opt.source,settings, opt.rowClass, opt.callback);
        }
    }

    $.fn.updateFromObj = function(obj,settings,rowClass, callback){
        settings.table.find("tbody").empty();
        $.each(obj, function(i,line) {
            var tableRow = $("<tr></tr>").addClass(rowClass);

            $.each(settings.json, function(j, identity) {
                if(identity == '*') {
                    tableRow.append($("<td>"+(i+1)+"</td>"));
                }
                else {
                    tableRow.append($("<td>" + line[this] + "</td>"));
                }
            });
            settings.table.append(tableRow);
        });


        if (typeof callback === "function") {
            callback();
        }

        $(window).trigger('resize');
    }

}( jQuery ));



</script>
<% writeCurrencyConverterOpeningScript %>

<style>
.colName
{
	display:none;
}

.colNumber
{
	display:none;
}

.colTeam
{
	display:none;
}

.jerseyStyleOptions
{
	background-color: red;
	border: dashed 1pt black;
	padding: 1pt 1pt 1pt 1pt;
}

.jerseyOptions
{
	text-align: left;
	border: solid 1pt black;
	background-color: #FFFFFF;
}

.jerseyTitle
{
	background-color : #A8A396;
	font-weight: bold;
}

#jerseyColor
{

}

</style>
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

		<form method="post" name="<%= MakeFormNameSafe(txtProdId) %>" action="http://dev.gamewearnow.com/addproduct.asp" onSubmit="return validateForm(this);">

		<!-- <form method="post" name="<%= MakeFormNameSafe(txtProdId) %>" action="<%= C_HomePath %>addproduct.asp" onSubmit="return validateForm(this);"> -->
		<input TYPE="hidden" NAME="PRODUCT_ID" VALUE="<%= txtProdId %>">
		<table border="0" width="100%" class="tdContent2" cellpadding="2" cellspacing="0">
		  <tr>
			<td align="center" valign="top"><%  '= detailImageOut %></td>
			<td align="left" valign="top">
				<h1 class="productName"><%= getProductInfo(txtProdId, enProduct_Name) %></h1>
				<% If getProductInfo(txtProdId, enProduct_MfgID) <> 1 Then %>
					<strong><%= C_ManufacturerNameS %>:</strong>&nbsp;<a href="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "URL", True) %>" title="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "MetaTitle", True) %>"><%= getMfgVendItem(getProductInfo(txtProdId, enProduct_MfgID), "Name", True) %></a><br />
				<% End If %>
				<% If getProductInfo(txtProdId, enProduct_VendorID) <> 1 Then %>
					<strong><%= C_VendorNameS %>:</strong>&nbsp;<a href="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "URL", False) %>" title="<%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "MetaTitle", False) %>"><%= getMfgVendItem(getProductInfo(txtProdId, enProduct_VendorID), "Name", False) %></a><br />
				<% End If %>
				<%= mstrProductDescription %><br />
				<%
				'<strong>%= C_Description %:</strong>&nbsp;
				'<strong> C_Price :</strong>&nbsp;
				'If iConverion = 1 Then
				'	If getProductInfo(txtProdId, enProduct_SaleIsActive) Then
				'		Response.Write "<span class=""itemOnSalePrice""><script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_Price)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
				'		Response.Write "<span class=""SalesPrice"">" & C_SPrice & ": <script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_SalePrice)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_SalePrice)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
				'		Response.Write "<span class=""YouSave"">" & C_YSave & ": <script>document.write(""" & FormatCurrency(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & " = ("" + OANDAconvert(" & trim(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></span><br />"
				'	Else
				'		Response.Write "<script>document.write(""" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & " = ("" + OANDAconvert(" & trim(getProductInfo(txtProdId, enProduct_Price)) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"
				'	End If
				'Else
				'	If getProductInfo(txtProdId, enProduct_SaleIsActive) Then
				'		Response.Write "<span class=""itemOnSalePrice"">" & FormatCurrency(getProductInfo(txtProdId, enProduct_Price)) & "</span><br />"
				'		Response.Write "<span class=""SalesPrice"">" & C_SPrice & ": " & FormatCurrency(getProductInfo(txtProdId, enProduct_SalePrice)) & "</span><br />"
				'		Response.Write "<span class=""YouSave"">" & C_YSave & ": " & FormatCurrency(CDbl(getProductInfo(txtProdId, enProduct_Price))-CDbl(getProductInfo(txtProdId, enProduct_SalePrice))) & "</span><br />"
				'	Else
				'		Response.Write FormatCurrency(getProductInfo(txtProdId, enProduct_Price))
				'	End If
				'End If	'iConverion = 1

				If cblnSF5AE Then
					SearchResults_GetProductInventory txtProdId
					'SearchResults_ShowMTPricesLink txtProdId

					'Response.Write "<hr />"
					If hasMTP(txtProdId) Then
						'Response.Write "<table border=0 cellpadding=2 cellspacing=0>"
						'Response.Write "<tr><th>Pricing</th></tr>"
						'Response.Write "<tr><td>"
						'Response.Write WriteMTPrices(getProductInfo(txtProdId, enProduct_MTP), getProductInfo(txtProdId, enProduct_Price))
						'Response.Write "</td></tr>"
						'Response.Write "</table>"

						Response.Write "<div align=""center"">"
						Response.Write WriteMTPricingTable(txtProdId, "Price Per Jersey (Unlettered)")
						Response.Write "</div>"

					End If
				End If
				%>
		      </td>
		    </tr>
		  </table>
		</form>
		<%= WriteJavaScript(mstrssAttributeExtenderjsOut) %>
<div align="left">
<!--#include file="detail_ReversibleDisplay.asp"-->

<form method="post" name="frmDetail" ID="frmDetail" action="<%= Session("DomainPath") %>addproduct.asp" onsubmit="return ValidateForm_Jersey(this);">
<input type="hidden" name="ssMPOPage" ID="ssMPOPage" value="1">
<input type="hidden" name="PRODUCT_ID" ID="PRODUCT_ID" value="<%= txtProdId %>">

<div id="jerseyOrder" class="jerseyOptions">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Arial; font-weight:bold" width="588">
	<tr>
		<td colspan="5" bgcolor="#A8A396"><b>Input Individual Jersey Information</b></td>
	</tr>
	<tr>
	<th>Player</th>
	<th width="15%">Size</th>
	<th width="15%" class="colNumber">Number</th>
	<th width="25%" class="colName">Name on Jersey</th>
	<th width="35%">Quantity of Jerseys for this Player</th>
	</tr>
	<%
	For EntryCounter = 1 To numEntries
		If cblnDebugJersey Then Response.Write "<fieldset><legend>Hidden Entry " & EntryCounter & "</legend>"
		For jerseyAttributeCounter = 0 To UBound(maryJerseyAttributes)
			If cblnDebugJersey Then
				'If maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Type) = 0 Or maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Type) = 2 Then
				'	Response.Write maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Name) & ": <input type=""text"" name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """><br />" & vbcrlf
				'End If
				Select Case maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Type)
					Case 0:	'hidden
						Response.Write maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Name) & ": "
						Response.Write "<input type=""text"" name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """>" & vbcrlf
						Response.Write " (name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """)"
						Response.Write "<br />" & vbcrlf
					Case 2:	'hidden text
						Response.Write maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Name) & ": "
						Response.Write "<input type=""text"" name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """ ID=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """ value="""">" & vbcrlf
						Response.Write " (name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """)"
						Response.Write "<br />" & vbcrlf
				End Select
			Else
				Select Case maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Type)
					Case 0:	'hidden
						Response.Write "<input type=""hidden"" name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """>" & vbcrlf
					Case 2:	'hidden text
						Response.Write "<input type=""hidden"" name=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """ ID=""attr" & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & """ value="""">" & vbcrlf
				End Select
			End If

'			response.Write "Attribute (" & i & "): " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Name) & " - " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Position) & " - " & maryJerseyAttributes(jerseyAttributeCounter)(enJerseyAttribute_Default) & "<br />"
		Next 'jerseyAttributeCounter
		If cblnDebugJersey Then Response.Write "</fieldset>"
	%>
	<tr>
	<td align="left"><b><font size="1">Player <%= EntryCounter %></font></b></td>
	<td align="center">
		<select name="attr<%= maryJerseyAttributes(enAttrPos_Size)(enJerseyAttribute_Position) & mstrPrefix %>" ID="attr<%= maryJerseyAttributes(enAttrPos_Size)(enJerseyAttribute_Position) & mstrPrefix %>" size="1" style="font-family: Arial Rounded MT Bold">
		<%
			attributeDetail = maryAttributes(maryJerseyAttributes(enAttrPos_Size)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
			For i = 0 To UBound(attributeDetail)
				If attributeDetail(i)(enAttributeDetail_Default) Then
					Response.Write "<option value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ selected>" & attributeDetail(i)(enAttributeDetail_Name) & "</option>"
				Else
					Response.Write "<option value=""" & attributeDetail(i)(enAttributeDetail_ID) & """>" & attributeDetail(i)(enAttributeDetail_Name) & "</option>"
				End If
			Next 'i
		%>
		</select>
	</td>
	<td align="center" class="colNumber">
		<input type="text" name="<%= "attr" & maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Default) %>" id="<%= "attr" & maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Default) %>" title="<%= maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Name) %>" value="" size="4"></input>
	</td>
	<td align="center" class="colName">
		<input type="text" name="<%= "attr" & maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Default) %>" id="<%= "attr" & maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Default) %>" title="<%= maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Name) %>" value="" size="15"></input>
	</td>
	<td align="center">
		<input style=""  type="text" name="QUANTITY" ID="QUANTITY" title="Quantity" size="3" onblur="return isInteger(this, true, 'Please enter an integer greater than one for the quantity')">
		<!--<input type="image" name="AddProduct" border="0" src="images/buttons/addtocart3.gif" alt="Add To Cart" />-->
	</td>
	</tr>
	<% Next 'EntryCounter %>
	<!--
	<tr>
	  <td colspan="5" align="center"><label for="customLogo"><% '= getProductInfo(cstrCustomLogo_ProductID, enProduct_Name) %></label></td>
	</tr>
	-->
	<tr>
	<td colspan="5" align="center"><input type="image" name="AddProduct" border="0" src="images/buttons/addtocart3.gif" alt="Add To Cart" /></td>
	</tr>

</table>

<div id="json_table"></div>

</div><input type="hidden" name="QUANTITY.<%= cstrCustomLogo_ProductID %>" id="customLogo" value="" />
</form>
</div>

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

<script>
$(document).ready(function() {
	//SET THE PRODUCT COLOR
	var enAttrPos_JerseyColor = '<%=response.write(request.form("enAttrPos_JerseyColor"))%>';
	$("td.jerseyDisplay:contains('"+enAttrPos_JerseyColor+"')").next().find('input').prop("checked", true).click();
	//SHOW AVLIBLE LETTING OPTIONS BASED ON REVERSE PRINTING
	var optionXa = '<%=response.write(request.form("optionXa"))%>';
	if(optionXa == "yes") {
		$(".jerseyDisplay2:contains('a:')").parent().prev().find('input').hide();
	}
	else {
		$(".jerseyDisplay2:contains('a:')").next().find('input').hide();
	};
	//SELECT THE LETTERING OPTION
	var letteringOption = '<%=response.write(request.form("letteringOption"))%>';
	$(".jerseyDisplay2 [type='radio']").filter(':visible').eq(letteringOption).prop("checked", true).click();
	//SIDE COLORS
	var colorOne = '<%=response.write(request.form("colorOne"))%>';
	var colorTwo = '<%=response.write(request.form("colorTwo"))%>';
	var sideOneColor = '<%=response.write(request.form("sideOneColor"))%>';
	var sideTwoColor = '<%=response.write(request.form("sideTwoColor"))%>';
	//one side
	if(colorTwo == ""){
		var colorColumn = $("#jerseyLetteringColor table tr td span:contains('"+colorOne+"')").attr('id');
		if(colorColumn == "spSide1Color") {
			var sideOneIndex = $("#spSide1Color:contains('"+colorOne+"')").closest('td').index();
			$("span:contains('"+colorOne+"')").closest('td').closest("table").find("tr td:nth-child("+(sideOneIndex)+")").each(function(){
				if($(this).text() == sideOneColor){
					$(this).next().find('input').prop("checked", true).click();
				}
			});
		}
		else {
			var sideTwoIndex = $("#spSide2Color:contains('"+colorOne+"')").closest('td').index();
			$("span:contains('"+colorOne+"')").closest('td').closest("table").find("tr td:nth-child("+(sideTwoIndex-2)+")").each(function(){
				if($(this).text() == sideOneColor){
					$(this).next().next().next().find('input').prop("checked", true).click();
				}
			});
		};
	}
	//two sides
	else {
		var sideOneIndex = $("#spSide1Color:contains('"+colorOne+"')").closest('td').index();
		$("span:contains('"+colorOne+"')").closest('td').closest("table").find("tr td:nth-child("+(sideOneIndex)+")").each(function(){
			if($(this).text() == sideOneColor){
				$(this).next().find('input').prop("checked", true).click();
			}
		});
		var sideTwoIndex = $("#spSide2Color:contains('"+colorTwo+"')").closest('td').index();
		$("span:contains('"+colorTwo+"')").closest('td').closest("table").find("tr td:nth-child("+(sideTwoIndex-2)+")").each(function(){
			if($(this).text() == sideTwoColor){
				$(this).next().next().next().find('input').prop("checked", true).click();
			}
		});
	};
	//FONT
	var font = '<%=response.write(request.form("font"))%>';
	$('#jerseyLetteringFont table').find("tr td:nth-child(1)").each(function() {
		if($(this).text() == font){
			$(this).next().find('input').prop("checked", true).click();
		}
	});
	//TEAM NAME
	var teamName = '<%=response.write(request.form("teamName"))%>';
	$(".jerseyDisplay:contains('Team Name: ')").find('input').val(teamName).change();
	//PLACEMENT
	var placement = '<%=response.write(request.form("placement"))%>';
	if(placement != "") {
		$(".jerseyDisplay:contains('Location')").find('select').find("option:contains('"+placement+"')").attr('selected', true).change();
	}
	else {
		$(".jerseyDisplay:contains('Location')").find('select').find("option:contains('Not Applicable')").attr('selected', true).change();
	};
	//PLAYER NAME LETTERING STYLE
	var playerLetteringStyle = '<%=response.write(request.form("playerLetteringStyle"))%>';
	$('#jerseyPlayerOptions table').find("tr td:nth-child(1)").each(function() {
		if($(this).text() == playerLetteringStyle){
			$(this).next().find('input').prop("checked", true).click();
		}
	});
	//TEAM NAME DESIGN
	var nameDesign = '<%=response.write(request.form("nameDesign"))%>';
	var nameStyle = '<%=response.write(request.form("nameDesignStyle"))%>';
	var graphic = '<%=response.write(request.form("graphic"))%>';
	var logo = '<%=response.write(request.form("logo"))%>';
	//team name style
	if(nameStyle != "") {
		$("td.jerseyTitle:contains('"+nameDesign+"')").closest("table").find("tr:gt(0):lt(4) td:nth-child(1)").each(function() {
			if($(this).text() == nameStyle){
				$(this).next().find('input').prop("checked", true).click();
			}
		});
	};
	//graphic
	if(graphic != "") {
		$("td.jerseyTitle:contains('"+nameDesign+"')").closest("table").find("tr:gt(5) td:nth-child(1)").each(function() {
			if($(this).text() == graphic){
				$(this).next().find('input').prop("checked", true).click();
			}
		});
	};
	//custom logo
	if(logo != "") {
		$("td.jerseyTitle:contains('"+nameDesign+"')").closest("table").find("tr td:nth-child(1)").each(function() {
			if($(this).text().indexOf(logo) != -1){
				$(this).next().find('input').prop("checked", true).click();
			}
		});
	};
	//JERSEY DETAILS
	var rows = '<%=response.write(request.form("jerseyRows"))%>';
	var json_source = '<%=response.write(request.form("json_source"))%>';
	var data = JSON.parse(json_source);
	var options = { source: data, };
	var detailsTable = $("<br><table id='row_details'></table>");
		detailsTable.jsonTable({
			head : ['#', 'Size', 'Price', 'Num', 'Name', 'Qty'],
			json : ['#', 'Size', 'Price', 'Num', 'Name', 'Qty']
		});
	detailsTable.jsonTableUpdate(options);

	$("#json_table").append(detailsTable);
	$('#row_details tr:eq(0)').remove(); //removes table header

  //run through each row
  function populateRow(counter, j_size, j_number, j_name, j_qty) {
  	$("#jerseyOrder tr:eq("+counter+")").find("select option").filter(function () { return $(this).html() == j_size; }).prop('selected', true)
		$("#jerseyOrder tr:eq("+counter+")").find('input[title*="Player Number"]').val(j_number)
		$("#jerseyOrder tr:eq("+counter+")").find('input[title*="Player Name"]').val(j_name);
		$("#jerseyOrder tr:eq("+counter+")").find('input[title*="Quantity"]').val(j_qty);
  };

	var counter = 2 //starts the row count after the table headers
  $('#row_details').find('tr').each(function() {
    var j_size = $(this).find('td:eq(1)').text();
    var j_number = $(this).find('td:eq(3)').text();
    var j_name = $(this).find('td:eq(4)').text();
    var j_qty = $(this).find('td:eq(5)').text();
    populateRow(counter, j_size, j_number, j_name, j_qty);
    counter ++
  });

});
</script>

</body>
</html>
<%
	Call cleanup_dbconnopen
%>