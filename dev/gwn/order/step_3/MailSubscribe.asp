<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.2

'@FILENAME: mailsubscribe.asp

'

'@DESCRIPTION: Allows Customer to Subscribe to Merchant Mailing List

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
<%
Dim sFirstName, sLastName,sPassword,sConfirmPass,sEmail
Dim mstrErrorMessage

	sEmail = Trim(Request.Form("emailadd"))
	sFirstName = Trim(Request.Form("fname"))
	sLastName = Trim(Request.Form("lname"))
	sPassword = Trim(Request.Form("password"))
	sConfirmPass = Trim(Request.Form("password2"))

	Set mclsCustomer = New clsCustomer
	With mclsCustomer
		If .LoadCustomerByEmailPassword(sEmail, sPassword) Then
			If .IsSubscribed Then
				mstrErrorMessage = "You are already subscribed. Thank you for your interest."
			Else
				If .SetSubscribed(True) Then
					mstrErrorMessage = sEmail & " has been added to our mailing list."
				Else
					mstrErrorMessage = "There was an error adding " & sEmail & " to our mailing list."
				End If
			End If
		Else
			'this means no record but may still be valid info
			mstrErrorMessage = .Message
			If Len(mstrErrorMessage) = 0 Then
				.custEmail = sEmail
				.custFirstName = sFirstName
				.custLastName = sLastName
				If Len(sPassword) = 0 Then
					.custPasswd = generatePassword
				Else
					.custPasswd = sPassword
				End If

				If .AddCustomer Then
					mstrErrorMessage = sEmail & " has been added to our mailing list."
				Else
					mstrErrorMessage = "There was an error adding " & sEmail & " to our mailing list."
					mstrErrorMessage = "There was an error adding " & sEmail & " to our mailing list.<br />" & .Message
				End If
			End If
		End If
	End With
	Set mclsCustomer = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> - Join Mailing List</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="no-cache">
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
  });

<!--

function validateForm(theForm)
{

    if (theForm.emailadd.value == "")
    {
		alert("Must Fill in Email Address");
		theForm.emailadd.focus()
		return false;
    }

    /*	Optional parameters
    if (theForm.fname.value == "")
    {
		alert("Must Fill in User Name");
		theForm.fname.focus()
		return false;
    }
    if (theForm.lName.value == "")
    {
		alert("Must Fill in Last Name");
		theForm.lName.focus()
		return false;
    }

    if (theForm.password.value == "")
    {
		alert("Must supply a password");
		theForm.password.focus()
		return false;
    }
    */

    if (theForm.password.value != "")
    {
		if (theForm.password2.value == "")
		{
			alert("Must supply a matching password confirmation");
			theForm.password2.focus()
			return false;
		}
		if (theForm.password.value != theForm.password2.value)
		{
			alert("Password and Password Confirmation Did Not Match");
			theForm.password.focus()
			return false;
		}
    }

}
//-->
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
#divCategoryTrail {
  margin-bottom: 15px;
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
.tdContent3 {
  background-image: none;
}
.inputImage {
  padding: 10px;
  margin-top: 0;
}
.myAccount {
  margin-top: -5px;
}
input {
  margin: 10px;
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
    </div>
  </div>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<% Call ShowMyAccountBreadCrumbsTrail("Mailing List", False) %>
<form method="post" action="mailsubscribe.asp" id="frmMailSubscribe" name="frmMailSubscribe" onsubmit="return validateForm(this)">
	<table class="Section" border="1" cellpadding="0" cellspacing="0">
      <tr>
        <td class="tdContent" align="left">
            <table border="0" cellspacing="15" cellpadding="2">
              <tr>
	            <td align="center" class="tdMiddleTopBanner">Mailing List</td>
              </tr>
              <tr>
                <td class="tdBottomTopBanner" align="left">Complete
                  the form below to subscribe to our mailing list.&nbsp;
                  Subscribers will be able to receive store newsletters, sale
                  announcements and other mailings of interest.</td>
              </tr>
              <tr>
                <td class="tdContent2" align="left">
                  <table border="0" cellpadding="2" cellspacing="0">
                    <% If Len(mstrErrorMessage) > 0 Then %>
                    <tr>
                      <td class="tdContent3" align="left" valign="middle">&nbsp;</td>
                      <td class="tdContent3" align="left" valign="middle"><strong><%= mstrErrorMessage %></strong></td>
                    </tr>
		            <% End If %>
                    <tr>
                      <td align="right">First Name:</td>
                      <td><input type="text" value="<%= sFirstname %>" name="fname" size="40" class="formDesign" id="fname"></td>
                    </tr>
                    <tr>
                      <td align="right">Last Name:</td>
                      <td><input type="text" value="<%= slastname %>" name="lName" size="40" class="formDesign" id="lName"></td>
                    </tr>
                    <tr>
                      <td align="right">Password:</td>
                      <td><input type="password" value="<%= spassword %>" name="password" size="40" class="formDesign" id="password"></td>
                    </tr>
                    <tr>
                      <td align="right">Confirm Password:</td>
                      <td><input type="password" value="<%= sconfirmpass %>" name="password2" size="40" class="formDesign" id="password2"></td>
                    </tr>
                    <tr>
                      <td align="right" valign="top"><font color="#FF0000">*</font>E-Mail Address:</td>
                      <td><input type="text" value="<%= sEmail %>" name="emailadd" size="40" class="formDesign" id="emailadd"></td>
                    </tr>
                    <tr>
                      <td width="100%" align="center" valign="top" colspan="2"><input type="image" class="inputImage" name="Submit" id="Submit" src="<%= C_BTN18 %>"></td>
                    </tr>
                    <tr>
                      <td align="center" valign="top">&nbsp;</td>
                      <td><font color="#FF0000">*</font> - Required Field<br />All other fields are optional and used for personalization only</td>
                    </tr>
                  </table>
	          </td>
	        </tr>
          </table>
	          </td>
	        </tr>
          </table>
</form>
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
Call cleanup_dbconnopen
%>