<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<TITLE>Custom Jerseys...Quick GameWearNow</TITLE>
<META http-equiv="Content-Language" content="en-us">
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<%
 'added because of dynamic cart display
 Response.CacheControl = "no-cache"
 Response.AddHeader "Pragma", "no-cache"
 Response.Expires = -1
%>
<meta name="keywords" content="keywords">
<meta name="description" content="description">
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
<style type="text/css">
.style2 {
	color: #00314A;
	background-color: #DEDEDE;
	text-align: center;
	font-family: "AvantGarde Bk BT";
	font-variant: small-caps;
	font-weight: bold;
	font-size: large;
	text-transform: uppercase;
}
.style3 {
	background-image: inherit;
	text-align: center;
	font-family: Arial, Helvetica, sans-serif;
	font-size: large;
	font-variant: small-caps;
	color: #000000;
	font-weight: bold;
	vertical-align: middle;
	text-transform: none;
}
.style6 {
	background-image: inherit;
	text-align: left;
	font-family: "BankGothic Md BT";
	font-size: large;
	font-variant: small-caps;
	color: #000000;
	font-weight: bold;
	vertical-align: middle;
	text-transform: uppercase;
}
.style7 {
	font-family: Verdana;
	vertical-align: middle;
	color: #00314A;
	font-size: large;
	font-variant: small-caps;
	text-transform: uppercase;
	text-align: center;
	background-color: #DEDEDE;
}
</style>
</HEAD>

<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
	<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
	<tr>
		<td>
		<table width="100%" border="0" cellspacing="1" cellpadding="3" style="width: 801px">
		<tr>
			<td align="center" class="tdMiddleTopBanner" colspan="3">
			&nbsp;</td>
		</tr>
		<tr>
			<td class="style7" colspan="3">
			&nbsp;</td>
		</tr>
		<!--webbot bot="PurpleText" PREVIEW="Begin Optional Confirmation Message Display" -->
		<% Call WriteThankYouMessage %>
		<!--webbot bot="PurpleText" PREVIEW="End Optional Confirmation Message Display" -->
		<tr>
			<td class="style3" colspan="3">
			<table style="width: 100%">
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>
			<img alt="Welcome to GameWearNow" src="images/Main%20Graphic.gif" width="757" height="228"></td>
		</tr>
		<tr>
			<td class="style6" style="width: 33%">
			&nbsp;</td>
			<td class="style3" style="width: 33%">
			&nbsp;</td>
			<td class="style3" style="width: 34%">
			&nbsp;</td>
		</tr>
		<tr>
			<td class="style3" colspan="3">
			<img alt="3 Easy Ordering Steps" src="ssl/images/Steps%201.0%20v2.gif" width="757" height="228"></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1278036-1";
urchinTracker();
</script>
<script>
console.log('legacy')
</script>
</body>
</html>
<%
	Call cleanup_dbconnopen
%>