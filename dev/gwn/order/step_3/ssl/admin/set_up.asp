<%@ LANGUAGE="VBSCRIPT" %><%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.3.0.1
'
'   Author     :   LaGarde, Incorporated
'
'   Description:   This file gathers the global variables used to operate
'				   your web store and logs them to the admin table in your 
'				   storefront database.
'
'   Notes      :  There are no configurable elements in this file.
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file is protected under the United States
'   copyright laws as an unpublished work, and is confidential and
'   proprietary to LaGarde, Incorporated.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   LaGarde, Incorporated is expressely prohibited.
'
'   (c) Copyright 1998 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%><%
	Dim DSN_Name
	DSN_Name = Session("DSN_Name")
	'Response.Write DSN_Name
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
	'Connection.Open "driver={Microsoft Access Driver (*.mdb)};dbq=e:\inetpub\vsroot\lagarde\thenaturestore\_private\naturestore.mdb"
		
	If Request.QueryString("Update") = "1" Then

	EMAIL_MESSAGE = Replace(Request.Form("EMAIL_MESSAGE"),"'","''")
	EMAIL_SUBJECT = Replace(Request.Form("EMAIL_SUBJECT"),"'","''")
	
	SQLStmt = "UPDATE Admin SET TAX_COUNTRY = '" & Trim(Request("Tax_Country")) & "', "
	SQLStmt = SQLStmt & "TAX_STATE = '" & Trim(Request("Tax_State")) & "', "
	SQLStmt = SQLStmt & "STATE_TAX_AMOUNT = '" & Trim(Request("State_Tax_Amount")) & "', "
	SQLStmt = SQLStmt & "COUNTRY_TAX_AMOUNT = '" & Trim(Request("Country_Tax_Amount")) & "', "
	SQLStmt = SQLStmt & "DOMAIN_NAME = '" & Trim(Request("DOMAIN_NAME")) & "', "
	SQLStmt = SQLStmt & "MAIL_METHOD = '" & Trim(Request("MAIL_METHOD")) & "', "
	SQLStmt = SQLStmt & "MAIL_SERVER = '" & Trim(Request("MAIL_SERVER")) & "', "
	SQLStmt = SQLStmt & "PRIMARY_EMAIL = '" & Trim(Request("PRIMARY_EMAIL")) & "', "
	SQLStmt = SQLStmt & "SECONDARY_EMAIL = '" & Trim(Request("SECONDARY_EMAIL")) & "', "
	SQLStmt = SQLStmt & "EMAIL_SUBJECT = '" & Trim(EMAIL_SUBJECT) &"', "
	SQLStmt = SQLStmt & "EMAIL_MESSAGE = '" & EMAIL_MESSAGE & "', "
	SQLStmt = SQLStmt & "MAIL_CC = '" & Trim(Request("MAIL_CC")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_A = '" & Trim(Request("SHIPPING_A")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_B = '" & Trim(Request("SHIPPING_B")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_C = '" & Trim(Request("SHIPPING_C")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_D = '" & Trim(Request("SHIPPING_D")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_E = '" & Trim(Request("SHIPPING_E")) & "', "
	SQLStmt = SQLStmt & "SHIPPING_F = '" & Trim(Request("SHIPPING_F")) & "', "
	SQLStmt = SQLStmt & "SHIPA_AMOUNT = '" & Trim(Request("SHIPA_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPAB_AMOUNT = '" & Trim(Request("SHIPAB_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPBC_AMOUNT = '" & Trim(Request("SHIPBC_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPCD_AMOUNT = '" & Trim(Request("SHIPCD_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPDE_AMOUNT = '" & Trim(Request("SHIPDE_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPEF_AMOUNT = '" & Trim(Request("SHIPEF_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SHIPF_UP_AMOUNT = '" & Trim(Request("SHIPF_UP_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "SF_PATH = '" & Trim(Request("SF_PATH")) & "', "
	SQLStmt = SQLStmt & "SSL_PATH = '" & Trim(Request("SSL_Path")) & "', "
	SQLStmt = SQLStmt & "SPECIAL_SHIP_AMOUNT = '" & Trim(Request("SPECIAL_SHIP_AMOUNT")) & "', "
	SQLStmt = SQLStmt & "TRANSACTION_METHOD = '" & Trim(Request("TRANSACTION_METHOD")) & "', "
	SQLStmt = SQLStmt & "PAYMENT_SERVER = '" & Trim(Request("PAYMENT_SERVER")) & "',"
	SQLStmt = SQLStmt & "LOGIN = '" & Trim(Request("LOGIN")) & "', "	
	SQLStmt = SQLStmt & "MERCHANT_TYPE = '" & Trim(Request("MERCHANT_TYPE")) & "' "
	SQLStmt = SQLStmt & "WHERE ID = " & Request.QueryString("ID") & " "
	'Response.Write SQLStmt
	Set RSUpdt = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * FROM Admin"
	Set RSConfirmUpdt = Connection.Execute(SQLStmt)

	ElseIf Request.QueryString("Update") = "0" Then

	SQLStmt = "SELECT * FROM Admin"
	
	Set RS = Connection.Execute(SQLStmt)

	End If

	
%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title>New Store Set-Up</title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body style="font-family: Arial; font-size: 9 pt">
<% If Request.QueryString("UPDATE") = "1" Then %>
<table width="90%" cellspacing="2" cellpadding="2" border="0" align="center">
<tr>
    <td width="50%" colspan="2" bgcolor="#A09A8B"><p align="center"><strong><big>Store Configuration Confirmation</big></strong></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" bgcolor="#A09A8B">The store configuration has been
    updated as shown below.&nbsp; You may change these variables at any time by re-running the
    New Store Set-Up routine.&nbsp; </td>
  </tr>
  <tr>
    <td width="50%" align="right">Tax Country:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("Tax_Country") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Country Tax Rate:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("Country_Tax_Amount") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Tax State or Area:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("Tax_State") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Tax Rate:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("State_Tax_Amount") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Domain Name:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("DOMAIN_NAME") %></td>
  </tr>
 
  <tr>
    <td width="50%" align="right">Application Path:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("SF_PATH") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Secure Directory (SSL) Path:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("SSL_PATH") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Transaction Method:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("TRANSACTION_METHOD") %></td>
  </tr>
  <% If (RSConfirmUpdt("TRANSACTION_METHOD")) = "AuthorizeNet" Then %>
  <tr>
    <td colspan="2"><a href="authchange.htm"><center><b>Send StoreFront Configuration to AuthorizeNet</b></center></a></td>
  </tr>
  <% End If %>
  <tr>
    <td width="50%" align="right">Payment Server:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("PAYMENT_SERVER") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Merchant Login:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("LOGIN") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Merchant Type:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("MERCHANT_TYPE") %></td>
  </tr>
<tr>
   <td width="50%" align="right">EMail Method</td>
  <td width="50%" align="left"><%= RSConfirmUpdt("MAIL_METHOD") %>
</td>
</tr>
<tr>
   <td width="50%" align="right">Mail Server Address</td>
  <td width="50%" align="left"><%= RSConfirmUpdt("MAIL_SERVER") %>
</td>
</tr>


  <tr>
    <td width="50%" align="right">Primary E-Mail Recipient:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("PRIMARY_EMAIL") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Secondary E-Mail Recipient: </td>
    <td width="50%"><%= RSConfirmUpdt("SECONDARY_EMAIL") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Send Credit Card Data to Merchant via E-Mail:</td>
    <td width="50%"><% IF RSConfirmUpdt("MAIL_CC") = "on" Then %>Yes<% Else %>No<% End If %></td>
  </tr>
  <tr>
  </tr>
  <tr>
    <td width="50%" align="right">EMail Subject Line:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("EMAIL_SUBJECT") %></td>
  </tr>
  <tr>
    <td width="50%" align="right">EMail Message:&nbsp; </td>
    <td width="50%"><%= RSConfirmUpdt("EMAIL_MESSAGE") %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><br>
      <br>
      <p><u>Shipping Cost Schedule</u></p>
      <br>
    </td>
  </tr>
  <tr>
    <td width="50%">For sales amount totals up to: <%= FormatCurrency(RSConfirmUpdt("SHIPPING_A")) %></td>
    <td width="50%">The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPA_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales from the amount above up to :<%= FormatCurrency(RSConfirmUpdt("SHIPPING_B")) %><p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPAB_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales from the amount above up to :<%= FormatCurrency(RSConfirmUpdt("SHIPPING_C")) %><p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPBC_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales from the amount above up to :<%= FormatCurrency(RSConfirmUpdt("SHIPPING_D")) %><p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPCD_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales from the amount above up to :<%= FormatCurrency(RSConfirmUpdt("SHIPPING_E")) %><p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPDE_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales from the amount above up to :<%= FormatCurrency(RSConfirmUpdt("SHIPPING_F")) %><p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPEF_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" valign="top">Sales over the amount shown above:<p>&nbsp;</p>
    </td>
    <td width="50%" valign="top"> The shipping charge is: <%= FormatCurrency(RSConfirmUpdt("SHIPF_UP_AMOUNT")) %></td>
  </tr>
  <tr>
    <td width="50%" colspan="2" align="center"><hr>
      </td>
  </tr>
  <tr>
    <td width="50%" colspan="2">For 2nd Day Air add this amount to the amount
    calculated for standard Ground Shipment:&nbsp; <%= FormatCurrency(RSConfirmUpdt("SPECIAL_SHIP_AMOUNT")) %></td>
  </tr>
</table>
<% End If %><% If Request.QueryString("UPDATE") = "0" Then %>
<table width="90%" cellspacing="2" cellpadding="2" border="0" align="center">
 <tr>
    <td bgcolor="#A09A8B"><p align="center"><strong><big>Store Administration</big></strong></p>
    </td>
  </tr>
  <tr>
    <td bgcolor="#A09A8B"> <p>The new store set-up routine allows you to set
    global variables for managing your store.&nbsp; The shipping rates, tax state, tax rate,
    and e-mail variables will be applied to each order processed through your web store.&nbsp;
    Default values have been pre-loaded into your store configuration files and are shown
    below.&nbsp; Be sure to check each variable and set it for your particular application.</p>
    </td>
  </tr>
</table>


<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Tax_Country.selectedIndex < 0)
  {
    alert("Please select one of the \"Tax_Country\" options.");
    theForm.Tax_Country.focus();
    return (false);
  }

  if (theForm.Tax_State.selectedIndex < 0)
  {
    alert("Please select one of the \"Tax_State\" options.");
    theForm.Tax_State.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="set_up.asp?UPDATE=1&amp;ID=<%= RS("ID") %>" align="center" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
  <table border="0" cellpadding="2" cellspacing="2" width="90%" align="center">
    <tr>
      <td width="50%" colspan="2" bgcolor="#A09A8B">The first section of the
      set-up routine deals with establishing the taxable country and state.&nbsp; Remember, it is your
      responsibility to accurately determine all applicable laws pertaining to the operation of
      your store and to configure the store correctly to comply with those laws. </td>
    </tr>
    <tr>
      <td width="50%" align="right">Tax Country:&nbsp; </td>
      <td width="50%"><!--webbot bot="Validation" s-display-name="Tax_Country" b-value-required="TRUE" --> <select name="Tax_Country" size="1">
          <option selected><%= RS("Tax_COUNTRY") %></option>
          <option> United States
		  </option>
          <option> Afghanistan
		  </option>
          <option> Albania
		  </option>
          <option> Algeria
		  </option>
          <option> American Samoa
		  </option>
          <option> Andorra
		  </option>
          <option> Angola
		  </option>
          <option> Antarctica
		  </option>
          <option> Antigua &amp; Barbuda
		  </option>
          <option> Argentina
		  </option>
          <option> Armenia
		  </option>
          <option> Aruba
		  </option>
          <option> Australia
		  </option>
          <option> Austria
		  </option>
          <option> Azerbaijan
		  </option>
          <option> Bahamas
		  </option>
          <option> Bahrain
		  </option>
          <option> Bangladesh
		  </option>
          <option> Barbados
		  </option>
          <option value="Republic of Belarus"> Belarus
		  </option>
          <option> Belgium
		  </option>
          <option> Belize
		  </option>
          <option> Benin
		  </option>
          <option> Bermuda
		  </option>
          <option> Bhutan
		  </option>
          <option> Bolivia
		  </option>
          <option> Bosnia &amp; Herzegovina
		  </option>
          <option> Botswana
		  </option>
          <option> Bouvet Island
		  </option>
          <option> Brazil
		  </option>
          <option> British Antarctica Territory
		  </option>
          <option> British Indian Ocean Territory
		  </option>
          <option> British West Indies
		  </option>
          <option> Brunei
		  </option>
          <option> Bulgaria
		  </option>
          <option> Burundi
		  </option>
          <option> Caledonia
		  </option>
          <option> Cambodia
		  </option>
          <option> Cameroon
		  </option>
          <option> Canada
		  </option>
          <option> Canary Islands
		  </option>
          <option> Canton And Enderbury Islands
		  </option>
          <option> Cape Verdi Islands
		  </option>
          <option> Cayman Islands
		  </option>
          <option> Central African Republic
		  </option>
          <option> Chad
		  </option>
          <option value="Chan. Island UK"> Channel Islands UK
		  </option>
          <option> Chile
		  </option>
          <option> China
		  </option>
          <option> Christmas Island
		  </option>
          <option> Cocos (Keeling) Islands
		  </option>
          <option value="Columbia"> Colombia
		  </option>
          <option value="Comoro Islands"> Comoros
		  </option>
          <option value="Congo"> Congo, Republic of
		  </option>
          <option> Congo, Democratic Republic of
		  </option>
          <option> Cook Islands
		  </option>
          <option> Costa Rica
		  </option>
          <option> Croatia
		  </option>
          <option> Cuba
		  </option>
          <option> Curacao
		  </option>
          <option> Cyprus
		  </option>
          <option> Czech Republic
		  </option>
          <option> Dahomey
		  </option>
          <option> Dem. People's Republic of Korea
		  </option>
          <option> Dem. Republic of Vietnam
		  </option>
          <option> Denmark
		  </option>
          <option> Djibouti
		  </option>
          <option> Dominica
		  </option>
          <option> Dominican Republic
		  </option>
          <option> Dronning Muad Land - Antarctica
		  </option>
          <option> Ecuador
		  </option>
          <option> Egypt
		  </option>
          <option> El Salvador
		  </option>
          <option> England. UK
		  </option>
          <option> Equatorial Guinea
		  </option>
          <option> Estonia
		  </option>
          <option> Ethiopia
		  </option>
          <option> Faeroe Islands
		  </option>
          <option> Falkland Islands
		  </option>
          <option> Fiji
		  </option>
          <option> Finland
		  </option>
          <option> France
		  </option>
          <option> French Guiana
		  </option>
          <option> French Polynesia
		  </option>
          <option> French Southern &amp; Antarctica
		  </option>
          <option> French Territory of Afars &amp; Issas
		  </option>
          <option> French West Indies
		  </option>
          <option> Gabon
		  </option>
          <option> Gambia
		  </option>
          <option> Gaza
		  </option>
          <option> Germany
		  </option>
          <option> Ghana
		  </option>
          <option> Gibraltar
		  </option>
          <option> Georgia
		  </option>
          <option> Greece
		  </option>
          <option> Greenland
		  </option>
          <option> Guadeloupe
		  </option>
          <option> Guam
		  </option>
          <option> Guatemala
		  </option>
          <option> Guinea
		  </option>
          <option> Guinea Bissau
		  </option>
          <option> Guyana
		  </option>
          <option> Haiti
		  </option>
          <option> Heard &amp; McDonald Islands
		  </option>
          <option> Holland
		  </option>
          <option> Honduras
		  </option>
          <option> Hong Kong
		  </option>
          <option> Hungary
		  </option>
          <option> Iceland
		  </option>
          <option> Ifni
		  </option>
          <option> India
		  </option>
          <option> Indonesia
		  </option>
          <option> Iran
		  </option>
          <option> Iraq
		  </option>
          <option value="Iraq-Saudi Arania Neutral Zone"> Iraq-Saudi Arabia Neutral Zone
		  </option>
          <option> Ireland
		  </option>
          <option> Israel
		  </option>
          <option> Italy
		  </option>
          <option> Ivory Coast
		  </option>
          <option> Jamaica
		  </option>
          <option> Japan
		  </option>
          <option> Johnston Atoll
		  </option>
          <option> Jordan
		  </option>
          <option> Kazakhstan
		  </option>
          <option> Kenya
		  </option>
          <option> Kuwait
		  </option>
          <option> Kyrgyzstan
		  </option>
          <option> Laos
		  </option>
          <option> Latvia
		  </option>
          <option> Lebanon
		  </option>
          <option> Leeward Islands
		  </option>
          <option> Lesotho
		  </option>
          <option> Liberia
		  </option>
          <option> Libya
		  </option>
          <option> Liechtenstein
		  </option>
          <option> Lithuania
		  </option>
          <option> Luxembourg
		  </option>
          <option> Macau
		  </option>
          <option> Madagascar
		  </option>
          <option> Malawi
		  </option>
          <option> Malaysia
		  </option>
          <option> Maldives
		  </option>
          <option> Mali
		  </option>
          <option> Malta
		  </option>
          <option> Mariana Islands
		  </option>
          <option> Martinique
		  </option>
          <option> Mauritania
		  </option>
          <option> Mauritius
		  </option>
          <option> Melanesia
		  </option>
          <option> Mexico
		  </option>
          <option> Micronesia
		  </option>
          <option> Midway Islands
		  </option>
          <option> Moldova
		  </option>
          <option> Monaco
		  </option>
          <option> Mongolia
		  </option>
          <option> Montserrat
		  </option>
          <option> Montenegro
		  </option>
          <option> Morocco
		  </option>
          <option> Mozambique
		  </option>
          <option> Myanmar
		  </option>
          <option> Namibia
		  </option>
          <option> Nauru
		  </option>
          <option> Navassa Island
		  </option>
          <option> Nepal
		  </option>
          <option> Netherlands
		  </option>
          <option> Netherlands Antilles
		  </option>
          <option> Neutral Zone
		  </option>
          <option> New Hebrides
		  </option>
          <option> New Zealand
		  </option>
          <option> Nicaragua
		  </option>
          <option> Niger
		  </option>
          <option> Nigeria
		  </option>
          <option> Niue
		  </option>
          <option> Norfolk Island
		  </option>
          <option value="North Ireland, UK"> Northern Ireland, UK
		  </option>
          <option> North Korea
		  </option>
          <option> Norway
		  </option>
          <option> Oman
		  </option>
          <option> Pacific Island
		  </option>
          <option> Pakistan
		  </option>
          <option> Panama
		  </option>
          <option> Papua New Guinea
		  </option>
          <option> Paracel Islands
		  </option>
          <option> Paraguay
		  </option>
          <option> Peru
		  </option>
          <option> Philippines
		  </option>
          <option> Pitcairn
		  </option>
          <option> Poland
		  </option>
          <option> Polynesia
		  </option>
          <option> Portugal
		  </option>
          <option> Portuguese Guinea
		  </option>
          <option> Portuguese Timor
		  </option>
          <option> Principe &amp; Sao Tome
		  </option>
          <option> Puerto Rico
		  </option>
          <option> Qatar
		  </option>
          <option value="Congo, Republic of"> Republic of Congo
		  </option>
          <option value="South Korea"> Republic of Korea
		  </option>
          <option> Reunion
		  </option>
          <option> Romania
		  </option>
          <option> Russia
		  </option>
          <option> Rwanda
		  </option>
          <option> Ryukyu Islands
		  </option>
          <option> Sabah
		  </option>
          <option> San Marino
		  </option>
          <option value="Sao Tome"> Sao Tome &amp; Principe
		  </option>
          <option> Saudi Arabia
		  </option>
          <option> Scotland, UK
		  </option>
          <option> Senegal
		  </option>
          <option> Serbia
		  </option>
          <option> Seychelles
		  </option>
          <option> Sierra Leone
		  </option>
          <option> Sikkim
		  </option>
          <option> Singapore
		  </option>
          <option> Slovakia
		  </option>
          <option> Slovenia
		  </option>
          <option> Solomon Islands
		  </option>
          <option> Somalia
		  </option>
          <option> Somaliliand
		  </option>
          <option> South Africa
		  </option>
          <option> South Korea
		  </option>
          <option> Spain
		  </option>
          <option> Spanish Sahara
		  </option>
          <option> Spartly Islands
		  </option>
          <option> Sri Lanka
		  </option>
          <option> St. Christopher-Nevis-Anguilla
		  </option>
          <option> St. Helena
		  </option>
          <option> St. Kitts
		  </option>
          <option> St. Lucia
		  </option>
          <option> St. Pierre &amp; Miquelon
		  </option>
          <option> St. Vincent
		  </option>
          <option> Sudan
		  </option>
          <option> Surinam
		  </option>
          <option> Svalbard &amp; Jan Mayen Islands
		  </option>
          <option> Swaziland
		  </option>
          <option> Sweden
		  </option>
          <option> Switzerland
		  </option>
          <option> Syrian Arab Republic
		  </option>
          <option> Taiwan
		  </option>
          <option> Tanzania
		  </option>
          <option> Thailand
		  </option>
          <option> Togo
		  </option>
          <option> Tonga
		  </option>
          <option> Transkei
		  </option>
          <option> Trinidad/Tobago
		  </option>
          <option> Tunisia
		  </option>
          <option> Turkey
		  </option>
          <option> Turkmenistan
		  </option>
          <option> Turks &amp; Caicos Islands
		  </option>
          <option> Uganda
		  </option>
          <option> Ukraine
		  </option>
          <option> United Arab Emirates
		  </option>
          <option> United Kingdom
		  </option>
          <option> United States
		  </option>
          <option> Uruguay
		  </option>
          <option> US Pacific Island
		  </option>
          <option> US Virgin Islands
		  </option>
          <option> Uzbekistan
		  </option>
          <option> Vanuatu
		  </option>
          <option> Vatican City
		  </option>
          <option> Venezuela
		  </option>
          <option> Vietnam
		  </option>
          <option> Virgin Islands (British)
		  </option>
          <option value="US Virgin Islands"> Virgin Islands (US)
		  </option>
          <option> Wake Island
		  </option>
          <option> Wales, UK
		  </option>
          <option> West Indies
		  </option>
          <option> Western Samoa
		  </option>
          <option> Windward Islands
		  </option>
          <option value="Dem. People's Republic of Yemen"> Yemen
		  </option>
          <option> Zambia
		  </option>
          <option> Zimbabwe
          </option>
          <option>All Others </option>
        </select></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">Enter the tax rate applicable
      to orders being shipped to the country selected above.&nbsp; Tax should be entered as a
      decimal value such as .065 for 6.5%</td>
        </tr>
        <tr>
          <td width="50%" align="right">Country Tax Rate:&nbsp; </td>
          <td width="50%"><input type="text" name="Country_Tax_Amount" value="<%= RS("Country_Tax_Amount") %>" size="20"></td>
        </tr>
        <tr>
          <td width="50%" align="right">Tax State or Area:&nbsp; </td>
          <td width="50%"><!--webbot bot="Validation" s-display-name="Tax_State" b-value-required="TRUE" --> <select name="Tax_State" size="1">
              <option selected><%= RS("Tax_State") %></option>
              <option>Not Applicable </option>
              <option> Alabama </option>
              <option> Alaska </option>
              <option> Arizona </option>
              <option> Arkansas </option>
              <option> California </option>
              <option> Colorado </option>
              <option> Connecticut </option>
              <option> Delaware </option>
              <option> Florida </option>
              <option> Georgia </option>
              <option> Hawaii </option>
              <option> Idaho </option>
              <option> Illinois </option>
              <option> Indiana </option>
              <option> Iowa </option>
              <option> Kansas </option>
              <option> Kentucky </option>
              <option> Louisiana </option>
              <option> Maine </option>
              <option> Maryland </option>
              <option> Massachusetts </option>
              <option> Michigan </option>
              <option> Minnesota </option>
              <option> Mississippi </option>
              <option> Missouri </option>
              <option> Montana </option>
              <option> Nebraska </option>
              <option> Nevada </option>
              <option> New Hampshire </option>
              <option> New Jersey </option>
              <option> New Mexico </option>
              <option> New York </option>
              <option> North Carolina </option>
              <option> North Dakota </option>
              <option> Ohio </option>
              <option> Oklahoma </option>
              <option> Oregon </option>
              <option> Pennsylvania </option>
              <option> Rhode Island </option>
              <option> South Carolina </option>
              <option> South Dakota </option>
              <option> Tennessee </option>
              <option> Texas </option>
              <option> Utah </option>
              <option> Vermont </option>
              <option> Virginia </option>
              <option> Washington </option>
              <option> Washington D.C. </option>
              <option> West Virginia </option>
              <option> Wisconsin </option>
              <option> Wyoming </option>
              <option> Alaska</option>
              <option> Hawaii </option>
              <option> Puerto Rico </option>
              <option> Alberta </option>
              <option> British Columbia </option>
              <option> Manitoba </option>
              <option> New Brunswick </option>
              <option> New Foundland </option>
              <option> Northwest Territories </option>
              <option> Nova Scotia </option>
              <option> Ontario </option>
              <option> Prince Edward Island </option>
              <option> Quebec </option>
              <option> Saskatchewan </option>
              <option> Yukon Territory </option>
              <option>All Others </option>
            </select></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">Enter the tax rate applicable
      to orders being shipped to the state or area selected above.&nbsp; Tax should be entered
      as a decimal value such as .065 for 6.5%</td>
        </tr>
        <tr>
          <td width="50%" align="right">State Tax Rate:&nbsp; </td>
          <td width="50%"><input type="text" name="State_Tax_Amount" value="<%= RS("State_Tax_Amount") %>" size="20"></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">Enter the full path to the
          root of your website.  This is the path that users will be sent to if they try to add items 
          to a closed order.</td>
        </tr>
        <tr>
          <td width="25%" align="right">Domain Root Path:</td>
          <td width="25%" colspan="3"><input type="text" name="DOMAIN_NAME" size="40" value="<%= RS("DOMAIN_NAME") %>"></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">In the box below you will
      need to enter the application path to the addproduct.asp file which is used by your
      StoreFront web store to build your customers orders. The path will begin with
      &quot;http.&quot; If you are working on a local, development webserver then this path will
      be something like: &quot;http://machinename/webname/addproduct.asp.&quot; If you are
      published to a production webserver then this path will probably be something like:
      &quot;http://yourdomain.com/addproduct.asp.&quot; Enter this path in the box below:</td>
        </tr>
        <tr>
          <td width="25%" align="right">Application Path:</td>
          <td width="25%" colspan="3"><input type="text" name="SF_PATH" size="40" value="<%= RS("SF_PATH") %>"></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">This box is used to set the
      path to the process_order.asp file. If you are using SSL you will need to move the ssl directory
      to the secure directory set up by the web administrator. The first file called to start the secure check
      out process is the the file &quot;process_order.asp&quot; This box should contain the full
      path to the process_order.asp file, including the file name itself. Such as
      &quot;https://www.securedomain.com/securedirectory/ssl/process_order.asp&quot; If you are not
      running in SSL; such as during development, etc, then this path should point to the ssl directory where the 
      process_order.asp file is running. This is a required field:</td>
        </tr>
        <tr>
          <td width="25%" align="right">SSL Path:</td>
          <td width="25%" colspan="3"><input type="text" name="SSL_PATH" size="40" value="<%= RS("SSL_PATH") %>"></td>
        </tr>
        <tr>
          <td width="50%" colspan="2" bgcolor="#A09A8B">In the input boxes below you
      will establish the variables for handling orders for your web store.&nbsp; The store is
      configured to send all orders via e-mail as well as a copy to the customer as a receipt. You will need to specify the mail component being used on the web server and the e-Mail addresses where you wish to have orders sent.</td>
        </tr>
<tr>
<td width="50%" align="right">Mail Method&nbsp; </td>

<td width="50%">
<select name="MAIL_METHOD">
    <option <% If Trim(RS("MAIL_METHOD")) = "CDONTS Mail" then %>selected<% end if %> value="CDONTS Mail">CDONTS Mail
	              </option>
                  <option <% If Trim(RS("MAIL_METHOD")) = "ASP Mail" then %>selected<% end if %> value="ASP Mail">ASP Mail
	              </option>
                  <option <% If Trim(RS("MAIL_METHOD")) = "OCX Mail" then %>selected<% end if %> value="OCX Mail">OCX Mail
	              </option>
                  <option <% If Trim(RS("MAIL_METHOD")) = "J Mail" then %>selected<% end if %> value="J Mail">J Mail
	              </option>
                  <option <% If Trim(RS("MAIL_METHOD")) = "Bamboo Mail" then %>selected<% end if %> value="Bamboo Mail">Bamboo Mail
	              </option>
                  <option <% If TRim(RS("MAIL_METHOD")) = "Simple Mail" then %>selected<% end if %> value="Simple Mail">Simple Mail
                  </option>
                  <option <% If Trim(RS("MAIL_METHOD")) = "SimpleMail 2.0" then %>selected<% end if %> value="SimpleMail 2.0">SimpleMail 2.0
                  </option>
				  <option <% if Trim(RS("MAIL_METHOD")) = "AB Mail" then %>selected<% end if %> value="AB Mail">AB Mail
	              </option>
                </select> </td>
              </tr>
<tr><td width="50%" align="right">Mail Server Address</td>
<td width="50%&quot;"><input type="text" name="MAIL_SERVER" value="<%= RS("MAIL_SERVER") %>"><td>
</tr>
        <tr>
          <td width="50%" align="right">Primary E-Mail Recipient:&nbsp; </td>
          <td width="50%"><input type="text" name="PRIMARY_EMAIL" value="<%= RS("PRIMARY_EMAIL") %>" size="40"></td>
        </tr>
        <tr>
          <td width="50%" align="right">Secondary E-Mail Recipient: </td>
          <td width="50%"><input type="text" name="SECONDARY_EMAIL" value="<%= RS("SECONDARY_EMAIL") %>" size="40"></td>
        </tr>
        <tr>
          <td width="50%">Send Credit Card Data to Merchant via E-Mail: </td>
            <td width="25%"><input type="checkbox" name="MAIL_CC" <% If RS("MAIL_CC") = "on" Then %>checked <% Else %><% End If %>>
            </td>
            </tr>

            <tr>
              <td width="50%" align="right">E-Mail Subject Line: </td>
              <td width="50%"><input type="text" name="EMAIL_SUBJECT" value="<%= RS("EMAIL_SUBJECT") %>" size="40"></td>
            </tr>
            <tr>
              <td width="50%" colspan="2" bgcolor="#A09A8B">The text entered below will
      appear as a header in the e-mail message which is sent to the customer as a receipt.
      &nbsp; </td>
            </tr>
            <tr>
              <td width="50%" align="right">Mail Message: </td>
              <td width="50%"><textarea rows="2" name="EMAIL_MESSAGE" cols="60"><%= RS("EMAIL_MESSAGE") %></textarea></td>
            </tr>
            <tr>
              <td width="50%" colspan="2" bgcolor="#A09A8B">From the list below indicate
      how you are going to process orders. The default method is for EMail Merchant Processing
      where all orders are sent to the primary and secondary store adminsistrator for processing
      outside of the StoreFront system. You may also elect to use an on-line authorization and
      transaction processing service.&nbsp; StoreFront currently has integrated support for
      AuthorizeNet (www.authorizenet.com), CyberCash (www.cybercash.com), and CarmelCash
      (www.carmelww.com). You will need to establish an account directly with one of these
      transaction services before configuring this option. The default configuration is for
      EMail Merchant Processing</td>
            </tr>
            <tr>
              <td width="25%">Transaction Method</td>
              <td width="25%"> <select name="TRANSACTION_METHOD">
                  <option <% if rs("transaction_method") = "CyberCash 3.2" then %>selected<% end if %>>CyberCash 3.2
	              </option>
                  <option <% if rs("transaction_method") = "CyberCash 2.14" then %>selected<% end if %>>CyberCash 2.14
	              </option>
                  <option <% if rs("transaction_method") = "AuthorizeNet" then %>selected<% end if %>>AuthorizeNet
	              </option>
                  <option <% if rs("transaction_method") = "AuthorizeNet-Direct" then %>selected<% end if %>>AuthorizeNet-Direct
	              </option>
                  <option <% if rs("transaction_method") = "CarmelCash" then %>selected<% end if %>>CarmelCash
	              </option>
				  <option <% if rs("transaction_method") = "PCAuthorize" then %>selected<% end if %>>PC Authorize
	              </option>
                  <option <% if rs("transaction_method") = "EMail Merchant Processing" then %>selected<% end if %>>EMail Merchant Processing
                  </option>
                </select> </td>
              </tr>
              <tr>
                <td width="50%" colspan="2" bgcolor="#A09A8B">If you have set up to use one
      of the on-line transaction services listed above then you will need to complete these
      fields. This information will be provided by the processing service. Disregard this
      section if using EMail Merchant Processing.</td>
              </tr>
              <tr>
                <td width="25%">Payment Server Path</td>
                <td width="25%"><input type="text" name="PAYMENT_SERVER" size="60" value="<%= RS("PAYMENT_SERVER") %>"></td>
              </tr>
              <tr>
                <td width="25%">Payment Server Login</td>
                <td width="25%"><input type="text" name="LOGIN" size="20" value="<%= RS("LOGIN") %>"></td>
              </tr>
              <tr>
                <td width="25%" valign="top">Merchant Type</td>
                <td width="25%"> 
	              <input type="radio" <% if trim(rs("merchant_type")) = "normal_auth" then %> checked <% else %> <% end if %> name="MERCHANT_TYPE" value="normal_auth">Normal Authorization<br>
                  <input type="radio" <% if trim(rs("merchant_type")) = "authonly" then %> checked <% else %> <% end if %> name="MERCHANT_TYPE" value="authonly">AuthOnly<br>
                  <input type="radio" <% if trim(rs("merchant_type")) = "credit" then %> checked <% else %> <% end if %> name="MERCHANT_TYPE" value="credit">Credit<br>
                  <input type="radio" <% if trim(rs("merchant_type")) = "authcapture" then %> checked <% else %> <% end if %> name="MERCHANT_TYPE" value="authcapture">AuthCapture</td>
              </tr>
              <tr>
                <td width="50%" colspan="2" bgcolor="#A09A8B">The fields below are used to
      set up the shipping rates applicable to orders from your web store.&nbsp; The standard
      configuration for StoreFront bases charges on the total amount of the customer's
      order.&nbsp; The standard configuration is set up to allow for standard Ground Shipment
      and 2nd Day Air Shipping.&nbsp; The 2nd Day Air Shipping Rate is calculated as an amount
      added to the Standard Ground Shipment charges.&nbsp; The amount added is the same across
      all rate categories. </td>
              </tr>
              <tr>
                <td width="50%">For sales amount<br>
                  totals up to: <input type="text" name="SHIPPING_A" value="<%= FormatCurrency(RS("SHIPPING_A")) %>" size="6"></td>
                <td width="50%" valign="bottom">The shipping charge is: <input type="text" name="SHIPA_AMOUNT" value="<%= FormatCurrency(RS("SHIPA_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales from the amount<br>
                  above up to :<input type="text" name="SHIPPING_B" value="<%= FormatCurrency(RS("SHIPPING_B")) %>" size="6"><p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPAB_AMOUNT" value="<%= FormatCurrency(RS("SHIPAB_AMOUNT"))  %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales from the amount<br>
                  above up to :<input type="text" name="SHIPPING_C" value="<%= FormatCurrency(RS("SHIPPING_C")) %>" size="6"><p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPBC_AMOUNT" value="<%= FormatCurrency(RS("SHIPBC_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales from the amount<br>
                  above up to :<input type="text" name="SHIPPING_D" value="<%= FormatCurrency(RS("SHIPPING_D")) %>" size="6"><p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPCD_AMOUNT" value="<%= FormatCurrency(RS("SHIPCD_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales from the amount<br>
                  above up to :<input type="text" name="SHIPPING_E" value="<%= FormatCurrency(RS("SHIPPING_E")) %>" size="6"><p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPDE_AMOUNT" value="<%= FormatCurrency(RS("SHIPDE_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales from the amount<br>
                  above up to :<input type="text" name="SHIPPING_F" value="<%= FormatCurrency(RS("SHIPPING_F")) %>" size="6"><p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPEF_AMOUNT" value="<%= FormatCurrency(RS("SHIPEF_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" valign="top">Sales over the amount shown above:<p>&nbsp;</p>
                </td>
                <td width="50%" valign="top"><br>
      The shipping charge is: <input type="text" name="SHIPF_UP_AMOUNT" value="<%= FormatCurrency(RS("SHIPF_UP_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td width="50%" colspan="2">For 2nd Day Air add this amount to the amount
      calculated for standard Ground Shipment:&nbsp; <input type="text" name="SPECIAL_SHIP_AMOUNT" value="<%= FormatCurrency(RS("SPECIAL_SHIP_AMOUNT")) %>" size="6"></td>
              </tr>
              <tr>
                <td colspan="2" align="right"><input type="submit" name="Submit" value="Create StoreFront Web Store Set-Up"></td>
              </tr>
            </table>
          </form>

          <% End If %>
          <p>&nbsp; </p>
	<% Connection.Close %>
          <p align="center"><a href="prodadd.htm">Add Product</a> | <a href="proddelete.htm">Delete Product</a> | <a href="prodlist.asp">List Products</a> | <a href="prodedit.htm">Edit Product</a><br>
          <a href="reports.htm">Sales Reporting</a> | <a href="set_up.asp?Update=0">Store Set-Up</a></p>
        </body>
      </html>
