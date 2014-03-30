<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<%


'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.2

'@FILENAME: advancedsearch.asp

'

'@DESCRIPTION: Product Search Page

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
 dim sSubCategories
 dim FrontPage_Form1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Advanced Search Page</title>
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
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
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

/******************************************************************
   convert_date()

   Function to convert supplied dates to format - dd/mm/yyyy.
	Valid input dates =
		ddmmyy, ddmmmyy, ddmmyyyy, ddmmmyyyy,
		d/m/yy, dd/m/yy, d/mm/yy, dd/mm/yy, d/mmm/yy, dd/mmm/yy,
		d/m/yyyy, dd/m/yyyy, d/mm/yyyy, dd/mm/yyyy, d/mmm/yyyy, dd/mmm/yyyy
	Valid date seperators =
		'-','.','/',' ',':','_',','

	Calls convert_month()
			invalid_date()
			validate_date()
			validate_year()

   Author: Simon Kneafsey
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk
   Date Created: 4/9/00

   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/

function convert_date(field1)
{
var fLength = field1.value.length; // Length of supplied field in characters.
var divider_values = new Array ('-','.','/',' ',':','_',','); // Array to hold permitted date seperators.  Add in '\' value
var array_elements = 7; // Number of elements in the array - divider_values.
var day1 = new String(null); // day value holder
var month1 = new String(null); // month value holder
var year1 = new String(null); // year value holder
var divider1 = null; // divider holder
var outdate1 = null; // formatted date to send back to calling field holder
var counter1 = 0; // counter for divider looping
var divider_holder = new Array ('0','0','0'); // array to hold positions of dividers in dates
var s = String(field1.value); // supplied date value variable

//If field is empty do nothing
if ( fLength == 0 ) {
   return true;
}

// Deal with today or now
if ( field1.value.toUpperCase() == 'NOW' || field1.value.toUpperCase() == 'TODAY' ) {

	var newDate1 = new Date();

  		if (navigator.appName == "Netscape") {
    		var myYear1 = newDate1.getYear() + 1900;
  		}
  		else {
  			var myYear1 =newDate1.getYear();
  		}

	var myMonth1 = newDate1.getMonth()+1;
	var myDay1 = newDate1.getDate();
	field1.value = myDay1 + "/" + myMonth1 + "/" + myYear1;
	fLength = field1.value.length;//re-evaluate string length.
	s = String(field1.value)//re-evaluate the string value.
}

//Check the date is the required length
if ( fLength != 0 && (fLength < 6 || fLength > 11) ) {
	invalid_date(field1);
	return false;
	}

// Find position and type of divider in the date
for ( var i=0; i<3; i++ ) {
	for ( var x=0; x<array_elements; x++ ) {
		if ( s.indexOf(divider_values[x], counter1) != -1 ) {
			divider1 = divider_values[x];
			divider_holder[i] = s.indexOf(divider_values[x], counter1);
		   //alert(i + " divider1 = " + divider_holder[i]);
			counter1 = divider_holder[i] + 1;
			//alert(i + " counter1 = " + counter1);
			break;
		}
 	}
 }

// if element 2 is not 0 then more than 2 dividers have been found so date is invalid.
if ( divider_holder[2] != 0 ) {
   invalid_date(field1);
	return false;
}

// See if no dividers are present in the date string.
if ( divider_holder[0] == 0 && divider_holder[1] == 0 ) {

		//continue processing
		if ( fLength == 6 ) {//ddmmyy
   		//day1 = field1.value.substring(0,2);
     	//	month1 = field1.value.substring(2,4);
     	month1 = field1.value.substring(0,2);
     	day1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,6);
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false;
				}
			}

		else if ( fLength == 7 ) {//mmmddy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,5);
  		 month1= field1.value.substring(0,3);
  			day1 = field1.value.substring(3,5);
  			year1 = field1.value.substring(5,7);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false;
				}
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false;
				}
			}
		else if ( fLength == 8 ) {//mmddyyyy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,4);
  		 month1= field1.value.substring(0,2);
  			day1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,8);
			}
		else if ( fLength == 9 ) {//mmmddyyyy
   		//day1 = field1.value.substring(0,2);
  		//	month1 = field1.value.substring(2,5);
  		month1 = field1.value.substring(0,3);
  		day1 = field1.value.substring(3,5);
  			year1 = field1.value.substring(5,9);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false;
				}
			}

		if ( (outdate1 = validate_date(day1,month1,year1)) == false ) {
   		alert("The value " + field1.value + " is not a vaild date.\n\r" +
			"Please enter a valid date in the format mm/dd/yyyy");
			field1.focus();
			field1.select();
			return false;
			}

		field1.value = outdate1;
		return true;// All OK
		}

// 2 dividers are present so continue to process
if ( divider_holder[0] != 0 && divider_holder[1] != 0 ) {
  	//day1 = field1.value.substring(0, divider_holder[0]);
  	//month1 = field1.value.substring(divider_holder[0] + 1, divider_holder[1]);
  	 month1= field1.value.substring(0, divider_holder[0]);
  	day1 = field1.value.substring(divider_holder[0] + 1, divider_holder[1]);
  	//alert(month1);
  	year1 = field1.value.substring(divider_holder[1] + 1, field1.value.length);
	}

if ( isNaN(day1) && isNaN(year1) ) { // Check day and year are numeric
	invalid_date(field1);
	return false;
   }

if ( day1.length == 1 ) { //Make d day dd
   day1 = '0' + day1;
}

if ( month1.length == 1 ) {//Make m month mm
	month1 = '0' + month1;
}

if ( year1.length == 2 ) {//Make yy year yyyy
   if ( (year1 = validate_year(year1)) == false ) {
   	invalid_date(field1);
		return false;
		}
}

if ( month1.length == 3 || month1.length == 4 ) {//Make mmm month mm
   if ( (month1 = convert_month(month1)) == false) {
   	alert("month1" + month1);
   	invalid_date(field1);
   	return false;
   }
}

// Date components are OK
if ( (day1.length == 2 || month1.length == 2 || year1.length == 4) == false) {
   invalid_date(field1);
   return false;
}

//Validate the date
if ( (outdate1 = validate_date(day1, month1, year1)) == false ) {
   alert("The value " + field1.value + " is not a vaild date.\n\r" +
	"Please enter a valid date in the format mm/dd/yyyy");

	field1.focus();
	field1.select();

	return false;
}

// Redisplay the date in dd/mm/yyyy format
field1.value = outdate1;
return true;//All is well

}
/******************************************************************
   convert_month()

   Function to convert mmm month to mm month

   Called by convert_date()

   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk

   Notes:Please feel free to use/edit this script.  If you do please keep my comments and details
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function convert_month(monthIn) {

var month_values = new Array ("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC");

monthIn = monthIn.toUpperCase();

if ( monthIn.length == 3 ) {
	for ( var i=0; i<12; i++ )
		{
   	if ( monthIn == month_values[i] )
   		{
			monthIn = i + 1;
			if ( i != 10 && i != 11 && i != 12 )
				{
   			monthIn = '0' + monthIn;
				}
			return monthIn;
			}
		}
	}

else if ( monthIn.length == 4 && monthIn == 'SEPT') {
   monthIn = '09';
   return monthIn;
	}

else {
	return false;
	}
}
/******************************************************************
   invalid_date()

   If an entered date is deemed to be invalid, invali
   d_date() is called to display a warning message to
   the user.  Also returns focus to the date  in que
   stion and selects the date for edit.

   Called by convert_date()

   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk

   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function invalid_date(inField)
{
alert("The value " + inField.value + " is not in a vaild date format.\n\r" +
        "Please enter date in the format mm/dd/yyyy");
inField.focus();
inField.select();
return true
}
/******************************************************************
   validate_date()

   Validates date output from convert_date().  Checks
   day is valid for month, leap years, month !> 12,.

   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk

   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function validate_date(day2, month2, year2)
{
var DayArray = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var MonthArray = new Array("01","02","03","04","05","06","07","08","09","10","11","12");

var inpDate = month2 + day2 + year2;

var filter=/^[0-9]{2}[0-9]{2}[0-9]{4}$/;

//Check mmddyyyy date supplied
if (! filter.test(inpDate))
  {
  return false;
  }
/* Check Valid Month */
filter=/01|02|03|04|05|06|07|08|09|10|11|12/ ;
if (! filter.test(month2))
  {
  return false;
  }
/* Check For Leap Year */
var N = Number(year2);
if ( ( N%4==0 && N%100 !=0 ) || ( N%400==0 ) )
  	{
   DayArray[1]=29;
  	}
/* Check for valid days for month */
for(var ctr=0; ctr<=11; ctr++)
  	{
   if (MonthArray[ctr]==month2)
   	{
      if (day2<= DayArray[ctr] && day2 >0 )
        {
        inpDate = month2 + '/' + day2 + '/' + year2;
        return inpDate;
        }
      else
        {
        return false;
        }
   	}
   }

}
/******************************************************************
   validate_year()

   converts yy years to yyyy
   Uses a hinge date of 10
        < 10 = 20yy
        => 10 = 19yy.

   Called by convert_date() before validate_date().

   Author: Simon Kneafsey
   Date Created: 4/9/00
   Email: simonkneafsey@hotmail.com
   WebSite: www.simonkneafsey.co.uk

   Notes: Please feel free to use/edit this script.  If you do please keep my comments and details
   intact and notify me via a quick Email to the address above.  Enjoy!
*******************************************************************/
function validate_year(inYear)
{
if ( inYear < 10 )
	{
   inYear = "20" + inYear;
   return inYear;
	}
else if ( inYear >= 10 )
	{
   inYear = "19" + inYear;
   return inYear;
	}
else
	{
	return false;
	}
}
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
.inputImage {
  padding: 10px;
}
select {
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
<form method="get" action="search_results.asp">
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="15" cellpadding="3">
        <tr>
        <td align="center" class="tdMiddleTopBanner">Advanced Search</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner">Use the options below to perform a more selective search of our product database.  You can confine your search to only certain Manufacturers, Categories, or Vendors.  You can also choose to search only by items that have been added within a certain time range, or search within a specific price range.
            <tr>
              <td class="tdContent2">

                  <table border="0" width="100%" cellpadding="4">
                    <tr>
                      <td width="100%"><b>Enter Keyword(s):</b><br />
            &nbsp;&nbsp;&nbsp;&nbsp; <input class="formDesign" name="txtsearchParamTxt" size="40">
                        <p><b>Search using:</b><br />
            &nbsp;&nbsp;&nbsp;&nbsp; <select size="1" class="formDesign" name="txtsearchParamType">
                          <option selected value="ALL">All of the Keywords</option>
                          <option value="ANY">Any of the Keywords</option>
                          <option value="Exact">Exact Phrase</option></select></p>
                        <%If C_CategoryIsActive <> 0 Then%>
                        <p><b>Select a <%= C_CategoryNameS %>:</b><br />
            &nbsp;&nbsp;&nbsp;&nbsp;<% WriteSingleSelect %></p>
                        <%Else%>
                        <input type="hidden" name= "txtsearchParamCat" value="ALL">

			            <%End If
			            If C_MFGIsActive <> 0 Then%>
			            <p><b>Select a <%= C_ManufacturerNameS %>:</b><br />
			&nbsp;&nbsp;&nbsp;&nbsp; <select class="formDesign" size="1" name="txtsearchParamMan">
                          <option value="ALL">All <%= C_ManufacturerNameP %></option><%= getManufacturersList(0) %></select></p>
                        <%Else%>
                        <input type="hidden" name= "txtsearchParamMan" value="ALL">
                        <%End If%>
                        <%If C_VendorIsActive <> 0 Then%>
                        <p><b>Select a <%= C_VendorNameS %>:</b><br />
            &nbsp;&nbsp;&nbsp;&nbsp; <select class="formDesign" size="1" name="txtsearchParamVen">
                          <option value="ALL">All <%= C_VendorNameP %></option><%= getVendorList(0) %></select></p>
                        <%Else%>
                        <input type="hidden" name= "txtsearchParamVen" value="ALL">
                        <%End If%>
                        <%If C_AddedIsActive <> 0 Then%>
                        <p><b>Added to Inventory Between:</b><br />
            &nbsp;&nbsp;&nbsp;&nbsp;<input type="text" class="formDesign" name="txtDateAddedStart" size="8" onblur="javascript:convert_date(this)"> <b>And</b>
                         <input type="text" class="formDesign" name="txtDateAddedEnd" size="8" onblur="javascript:convert_date(this)"></p>
			            <%End If%>
			            <%If C_PriceIsActive <> 0 Then%>
			            <p><b>Price Between:</b><br />
			&nbsp;&nbsp;&nbsp;&nbsp;<input class="formDesign" type="text" name="txtPriceStart" size="8"> <b>To</b>
                         <input class="formDesign" type="text" name="txtPriceEnd" size="8"></p>
			            <%End If%>
			            <%If C_SaleIsActive <> 0 Then%>
			            <p><b>Show only sale items discounted at least</b>
			            &nbsp;<select class="formDesign" size="1" name="txtSale" id="txtSale">
			            <option value="" selected>All items</option>
			            <option value="0">All items on sale</option>
			            <option value="10">10%</option>
			            <option value="20">20%</option>
			            <option value="30">30%</option>
			            <option value="40">40%</option>
			            <option value="50">50%</option>
			            <option value="60">60%</option>
			            <option value="70">70%</option>
			            <option value="80">80%</option>
			            <option value="90">90%</option>
			            </select>
			            </p>
			            <%End If%>
			            <p align="center"><input type="image" class="inputImage" name="btnSearch" src="<%= C_BTN21 %>" alt="Search"></p>
			            </td>
                      </tr>
                    </table>
                    <input type="hidden" name="txtFromSearch" value="fromSearch">
                    <input type="hidden" name="iLevel" value="1">

	            </td>
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