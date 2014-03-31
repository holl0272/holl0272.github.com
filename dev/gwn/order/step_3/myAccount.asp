<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
'********************************************************************************
'*   myAccount Version SF 5.0		                                            *
'*   Release Version:	1.00.003                                                *
'*   Release Date:		September 29, 2002										*
'*   Revision Date:		September 30, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.00.003 (September 30, 2003)		                                *
'*   - Restructured add-on to work from root instead of myAccount folder		*
'*   - General clean-up															*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<!--#include file="SFLib/ssProductReview.asp"-->
<!--#include file="include_files/myAccount_CustomerInfoEditForm.asp"-->
<%

Dim mstrCallingPage
Dim mstrProblemReportID
Dim mstrMessage

	mstrAction = LoadRequestValue("Action")

	'Check for registration
	'Note: this check must be done BEFORE the cart total is displayed or it will not show the first pass through
	'get Certificate from querystring or form variables
	mstrCertificate = LoadRequestValue("Certificate")
	mstrProblemReportID = LoadRequestValue("problemReportID")

	If len(mstrCertificate) > 0 Then
		If mstrAction = "deleteGCRegistration" Then
			Call deleteGCRegisteredForUse(mstrCertificate)
		Else
			Call checkForCertificateEntry(mstrCertificate)
		End If
	ElseIf Len(session("ssGiftCertificate")) > 0 Then
		mstrGiftCertificateRegistrationMessage = "Certificate " & DisplayCertificateCodes(session("ssGiftCertificate")) & " will be credited during checkout."
	End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<title><%= C_STORENAME %>-Customer Account</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
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
    if(isMobile){
			$('input').on('focus', function(){
			  $('#footer').hide();
			}).on('blur', function(){
			  $('#footer').show();
			});
    };
		$("a[title='Return to home page']").attr('href', 'index.html');

		$(window).load(function() {
			$('#tdContent').css('opacity', 1);
			$('#tblMainContent').css('opacity', 1);
		});
	});
</script>

<style>
body {
	background-image: url('images/splash_bg.jpg');
	text-align: center;
}
#tdContent {
	margin: 0 auto 5%;
	opacity: 0;
}
#tblCategoryMenu, .tdTopBanner, .tdLeftNav {
	display: none;
}
.Section {
	width: 100%;
}
.tdContent2 {
	padding: 25px;
}
.tdContent2 table {
	margin-top: 10px;
}
.inputImage {
	padding: 10px;
	margin-top: 0;
}
.myAccount {
	margin-top: -5px;
}
.tbl {
	margin: 0 auto;
	background-color: white;
	width: 100%;
}
input {
	margin: 10px;
}
ul {
	list-style: none;
	margin-left: 10px;
	line-height: 2;
}
#footer ul {
	line-height: 1;
}
select#custState, select#custCountry {
	margin-left: 10px;
}
</style>

</head>

<body <%= mstrBodyStyle %>>

	<div id="header">
    <div id="gwn_logo">
      <a href="index.html" title="Home"><image src="images/gwn_logo.png" alt="GameWearNow Logo" style="margin-left: -25px;"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
        <br>
      <span class="title_txt" id="sub_title">MY ACCOUNT</span>
    </div>
  </div>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="15" cellpadding="0" id="tblMainContent" width="801px" style="opacity: 0;">
		<tr>
		<td>

<%
'************************************************************************************************************************
'
'	myAccount - Begin actual code
'
'	If you wish you can copy this section of code into your own .asp template file
'	Notes:
'			- Make sure you include all the include references at the top of this file
'			- Make sure you include the close database connection at the bottom of this file

'Response.Write "<H4>custID = " & custID & "</H4>"

On Error Goto 0
mstrCallingPage = LoadRequestValue("PrevPage")
If Len(mstrCallingPage) = 0 Then mstrCallingPage = Request.ServerVariables("SCRIPT_NAME")
If Not ((mstrAction = "createAccount") Or (mstrAction = "Create Account")) Then
	Call ProtectThisPage(mstrCallingPage)
End If

Select Case mstrAction
	Case "BuyersClubCreateRedemption":
		If ssBuyersClub_CreateRedemption(visitorLoggedInCustomerID) Then
			mstrMessage = mstrMessage_BuyersClub
			Call ShowMyAccountBreadCrumbsTrail("Redemptions", False)
			Call ShowMenu(False)
		Else
			Call ShowMyAccountBreadCrumbsTrail("Redemption Options", True)
			Call ShowBuyersClubRedemptionOptions(visitorLoggedInCustomerID)
		End If
	Case "ShowBuyersClubDetail":
		Call ShowMyAccountBreadCrumbsTrail("Earning History", False)
		Call ShowMenu(False)
		Call ShowBuyersClubHistory(visitorLoggedInCustomerID, True)
	Case "ViewBuyersClubRempdtion":
		Call ShowMyAccountBreadCrumbsTrail("Redemption History", False)
		Call ShowMenu(False)
		Call ShowBuyersClubHistory(visitorLoggedInCustomerID, True)
	Case "ShowBuyersClubRedemptionOptions":
		Call ShowMyAccountBreadCrumbsTrail("Redemption Options", True)
		Call ShowBuyersClubRedemptionOptions(visitorLoggedInCustomerID)
	Case "View","Update":
		Call ShowMyAccountBreadCrumbsTrail("My Profile", False)
		Call ProcessCustomerForm	'Note this function is in the include file CustomerInfoEditForm
		Call ShowMenu(False)
	Case "createAccount","Create Account":
		Call ShowMyAccountBreadCrumbsTrail("Create Account", False)
		Call ProcessCustomerForm	'Note this function is in the include file CustomerInfoEditForm
	Case "ChangePwd","EmailPwd","LogOff":
		'these are handled by ProtectThisPage
		If mblnShowMenu Then Call ShowMenu(True)
	Case "viewProblemReport":
		If isLoggedIn Then Call ShowMyAccountBreadCrumbsTrail("Problem Reports", False)
		If mblnShowMenu Then Call ShowMenu(True)
	Case Else
		If isLoggedIn Then Call ShowMyAccountBreadCrumbsTrail("", False)
		If mblnShowMenu Then Call ShowMenu(True)
End Select

'	myAccount - End actual code
'
'************************************************************************************************************************
%>
        </td>
      </tr>
    </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->

	<div id="footer">
    <ul id="horizontal-nav">
      <li class="not_selected"><a href="order.asp" title="Shopping Cart"><span><image src="../../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li id="current_page"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
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
Call cleanup_dbconnopen	'This line needs to be included to close database connection
%>