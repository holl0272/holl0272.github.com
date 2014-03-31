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
<link runat="server" rel="shortcut icon" type="image/png" href="favicon.ico">
<link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Lato:100,400,900|Josefin+Sans:100,400,700,400italic,700italic">
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

<style>
.black_overlay{
    opacity: 1 !important;
    display: block;
    position: fixed;
    top: 0%;
    left: 0%;
    width: 100%;
    height: 100%;
    z-index:1001;
    background-image: url('images/splash_bg.jpg');
    -moz-opacity: 0.8;
    opacity:.80;
    filter: alpha(opacity=80);
}
.white_content {
    opacity: 1 !important;;
    display: block;
    position: absolute;
    top: 25%;
    left: 25%;
    width: 50%;
    height: auto;
    padding: 16px;
    border: 16px solid #e8d606;
    background-color: #11013b;
    color: #cccdce;
    z-index:1002;
    overflow: auto;
    border-radius: 10px;
    text-align: center;
    font-size: 2em;
    font-weight: 900;
    font-family: 'Lato', sans-serif;
}

#fadingBarsG{
margin: 25px auto;
position:relative;
width:240px;
height:29px}

.fadingBarsG{
position:absolute;
top:0;
background-color:#e8d606;
width:29px;
height:29px;
-moz-animation-name:bounce_fadingBarsG;
-moz-animation-duration:1.7s;
-moz-animation-iteration-count:infinite;
-moz-animation-direction:linear;
-moz-transform:scale(.3);
-webkit-animation-name:bounce_fadingBarsG;
-webkit-animation-duration:1.7s;
-webkit-animation-iteration-count:infinite;
-webkit-animation-direction:linear;
-webkit-transform:scale(.3);
-ms-animation-name:bounce_fadingBarsG;
-ms-animation-duration:1.7s;
-ms-animation-iteration-count:infinite;
-ms-animation-direction:linear;
-ms-transform:scale(.3);
-o-animation-name:bounce_fadingBarsG;
-o-animation-duration:1.7s;
-o-animation-iteration-count:infinite;
-o-animation-direction:linear;
-o-transform:scale(.3);
animation-name:bounce_fadingBarsG;
animation-duration:1.7s;
animation-iteration-count:infinite;
animation-direction:linear;
transform:scale(.3);
}

#fadingBarsG_1{
left:0;
-moz-animation-delay:0.68s;
-webkit-animation-delay:0.68s;
-ms-animation-delay:0.68s;
-o-animation-delay:0.68s;
animation-delay:0.68s;
}

#fadingBarsG_2{
left:30px;
-moz-animation-delay:0.85s;
-webkit-animation-delay:0.85s;
-ms-animation-delay:0.85s;
-o-animation-delay:0.85s;
animation-delay:0.85s;
}

#fadingBarsG_3{
left:60px;
-moz-animation-delay:1.02s;
-webkit-animation-delay:1.02s;
-ms-animation-delay:1.02s;
-o-animation-delay:1.02s;
animation-delay:1.02s;
}

#fadingBarsG_4{
left:90px;
-moz-animation-delay:1.19s;
-webkit-animation-delay:1.19s;
-ms-animation-delay:1.19s;
-o-animation-delay:1.19s;
animation-delay:1.19s;
}

#fadingBarsG_5{
left:120px;
-moz-animation-delay:1.36s;
-webkit-animation-delay:1.36s;
-ms-animation-delay:1.36s;
-o-animation-delay:1.36s;
animation-delay:1.36s;
}

#fadingBarsG_6{
left:150px;
-moz-animation-delay:1.53s;
-webkit-animation-delay:1.53s;
-ms-animation-delay:1.53s;
-o-animation-delay:1.53s;
animation-delay:1.53s;
}

#fadingBarsG_7{
left:180px;
-moz-animation-delay:1.7s;
-webkit-animation-delay:1.7s;
-ms-animation-delay:1.7s;
-o-animation-delay:1.7s;
animation-delay:1.7s;
}

#fadingBarsG_8{
left:210px;
-moz-animation-delay:1.87s;
-webkit-animation-delay:1.87s;
-ms-animation-delay:1.87s;
-o-animation-delay:1.87s;
animation-delay:1.87s;
}

@-moz-keyframes bounce_fadingBarsG{
0%{
-moz-transform:scale(1);
background-color:#e8d606;
}

100%{
-moz-transform:scale(.3);
background-color:#11013b;
}

}

@-webkit-keyframes bounce_fadingBarsG{
0%{
-webkit-transform:scale(1);
background-color:#e8d606;
}

100%{
-webkit-transform:scale(.3);
background-color:#11013b;
}

}

@-ms-keyframes bounce_fadingBarsG{
0%{
-ms-transform:scale(1);
background-color:#e8d606;
}

100%{
-ms-transform:scale(.3);
background-color:#11013b;
}

}

@-o-keyframes bounce_fadingBarsG{
0%{
-o-transform:scale(1);
background-color:#e8d606;
}

100%{
-o-transform:scale(.3);
background-color:#11013b;
}

}

@keyframes bounce_fadingBarsG{
0%{
transform:scale(1);
background-color:#e8d606;
}

100%{
transform:scale(.3);
background-color:#11013b;
}

}
</style>


</HEAD>

<body <%= mstrBodyStyle %> >

<div id="light" class="white_content">
  <br>Thank you for visiting gamewearnow.com<br>We will be performing maintenance<br>Sunday evening, March 30th<br>We apologize for the inconvenience

  <div id="fadingBarsG">
    <div id="fadingBarsG_1" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_2" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_3" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_4" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_5" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_6" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_7" class="fadingBarsG">
    </div>
    <div id="fadingBarsG_8" class="fadingBarsG">
    </div>
  </div>
</div>
<div id="fade" class="black_overlay"></div>
<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->

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