<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<%
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4014.0.2

	'@FILENAME: search.asp




	'@DESCRIPTION: Search Page

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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Search Engine Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="keywords">
<meta name="description" content="description">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">
<link runat="server" rel="shortcut icon" type="image/png" href="favicon.ico">
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<link rel="stylesheet" href="css/main.css">
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/jquery-1.11.0.min.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
	$(document).ready(function() {

	var media = navigator.userAgent.toLowerCase();
	var isMobile = media.indexOf("mobile") > -1;
	if(isMobile) {
		$('#horizontal-nav li').css('padding-right', '10px');
	};

		WebFontConfig = {
		  google: { families: [ 'Lato:100,400,900:latin', 'Josefin+Sans:100,400,700,400italic,700italic:latin' ] }
		  };
		  (function() {
		    var wf = document.createElement('script');
		    wf.src = ('https:' == document.location.protocol ? 'https' : 'http') +
		      '://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js';
		    wf.type = 'text/javascript';
		    wf.async = 'true';
		    var s = document.getElementsByTagName('script')[0];
		    s.parentNode.insertBefore(wf, s);
		})();

		$('table.tdTopBanner').next().css('margin', '0 auto 10%');
		$('#frmPromo table').css('margin', '0 auto');

		$(".not_selected").hover(
		  function() {
		    $('#current_page a').css('color','#cccdce');
		  }, function() {
		    $('#current_page a').css('color','#e8d606');
		  }
		);
	});
</script>

<style>
body {
	background-image: url('images/splash_bg.jpg');
	text-align: center;
}
#tdContent {
	margin: 0 auto 5%;
}
#tblCategoryMenu, .tdTopBanner, .tdLeftNav {
	display: none;
}
</style>

</head>

<body <%= mstrBodyStyle %> onload="searchForm.txtsearchParamTxt.focus();">

	<div id="header">
    <div id="gwn_logo">
      <a href="index.html" title="Home"><image src="images/gwn_logo.png" alt="GameWearNow Logo" style="margin-left: -25px;"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
    </div>
  </div>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
	<tr>
		<td>
		<table width="100%" border="0" cellspacing="15" cellpadding="3">
			<tr>
			<td align="center" class="tdMiddleTopBanner">Search Store</td>
			</tr>
			<tr>
			<td class="tdBottomTopBanner2">
				Please input the word(s) that you would like to search for in our product database. For additional control you may choose to search on
				&quot;All Words&quot; or &quot;Any Words&quot; or for the &quot;Exact Phrase.&quot;&nbsp; For additional search options you may use our <i><a href="advancedsearch.asp"> Advanced Search</a></i> page.
			</td>
			</tr>
			<tr>
			<td class="tdContent2">
				<form method="get" name="searchForm" action="search_results.asp">
				<input type="hidden" name="iLevel" value="1">
				<input type="hidden" name="txtsearchParamMan" value="ALL">
				<input type="hidden" name="txtsearchParamVen" value="ALL">
				<input type="hidden" name="txtFromSearch" value="fromSearch">

				<table border="0" cellpadding="0" cellspacing="5" width="100%">
					<tr>
					<td align="center"><b>Search</b>&nbsp;&nbsp;<input type="text" class="formDesign" name="txtsearchParamTxt" size="20">&nbsp;&nbsp;<b>In</b>&nbsp;&nbsp;<% WriteSingleSelect %>&nbsp;<input type="image" class="inputImage" name="btnSearch" src="<%= C_BTN01 %>" alt="Search"></td>
					</tr>
					<tr>
					<td width="100%" align="center"><font class="Content_Small">
            			<input type="radio" value="ALL" name="txtsearchParamType" id="txtsearchParamType0" checked> <label for="txtsearchParamType0"><b>ALL</b> Words</label>
            			<input type="radio" name="txtsearchParamType" id="txtsearchParamType1" value="ANY"> <label for="txtsearchParamType1"><b>ANY</b> Words</label>
            			<input type="radio" name="txtsearchParamType" id="txtsearchParamType2" value="Exact"> <label for="txtsearchParamType2">Exact Phrase</label> |
						<a href="advancedsearch.asp"> Advanced Search</a></font>
					</td>
					</tr>
				</table>
				</form>
			</td>
			</tr>
          </table>
        </td>
      </tr>
    </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->

	<div id="footer">
    <ul id="horizontal-nav">
      <li class="not_selected"><a href="order.asp" title="Shopping Cart"><span><image src="../../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/privacy_policy/privacy_policy.html" title="Contact Us">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>
</body>
</html>
<%
Call cleanup_dbconnopen
%>