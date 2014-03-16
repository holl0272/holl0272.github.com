<%@ LANGUAGE="VBScript" %><%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.2.5.1
'
'   Author     :   LaGarde, Incorporated
'
'   Description:   Handles the creation of the product table.
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
	InPage = Request.ServerVariables("HTTP_REFERER")
	
	Description = Replace(Request.Form("Description"),"'","''")
	Message = Replace(Request.Form("Message"),"'","''")
	Manufacturer = Replace(Request.Form("Manufacturer"),"'","''")

	If Request("WEIGHT") ="" Then
	WEIGHT = "0"
	Else
	WEIGHT = Request("WEIGHT")
	End If

	Dim DSN_Name
	DSN_Name = Session("DSN_Name")

	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
	'Connection.Open "driver={Microsoft Access Driver (*.mdb)};dbq=e:\inetpub\vsroot\lagarde\thenaturestore\_private\naturestore.mdb"

	'Outputs the product detail for the requested product

	If Request.QueryString("Function") = "1" Then

	SQLStmt = "SELECT * FROM Product WHERE Product_ID = '" & Request("Product_ID") & "' "

	Set RS1 = Connection.Execute(SQLStmt)

	'Adds the product information to the products table.

	ElseIf Request("Function") = "2" Then
	MyCurrSymbol = FormatCurrency(1)
	MyCurrSymbol = Replace((MyCurrSymbol),"1","")
	MyCurrSymbol = Replace((MyCurrSymbol),",","")
	MyCurrSymbol = Replace((MyCurrSymbol),".","")
	MyCurrSymbol = Replace((MyCurrSymbol),"0","")

	ProdPrice = Replace(Request("Price"),MyCurrSymbol,"")	
	If MyCurrSymbol = "$" Then
	rPrice = Replace((ProdPrice),",","")
	Else
	rPrice = Replace((ProdPrice),",",".")
	End If
	
	SQLStmt = "INSERT INTO Product (Product_ID, Description, "
	SQLStmt = SQLStmt & "Price, Message, Category, Weight, Image_Path, Link, Mfg) "
	SQLStmt = SQLStmt & "VALUES ('" & Request("Product_ID") & "', "
	SQLStmt = SQLStmt & "'" & Description & "', "
	SQLStmt = SQLStmt & "'" & rPrice & "', "
	SQLStmt = SQLStmt & "'" & Message & "', '" & Request("Category") & "', "
	SQLStmt = SQLStmt & "'" & Weight & "', '" & Request("Image_Path") & "',"
	SQLStmt = SQLStmt & "'" & Request("Link") & "', '" & Manufacturer & "')"

	Set RS2 =Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * FROM Product WHERE PRODUCT_ID = '" & Request("Product_ID") & "' "

	Set RS2A = Connection.Execute(SQLStmt)

	'Deletes the selected product from the products table.

	ElseIf Request("Function") = "3" Then

	SQLStmt = "DELETE * FROM Product WHERE Product_ID = '" & Request("Product_ID") & "' "

	Set RS3 = Connection.Execute(SQLStmt)
	
	'Response.Redirect InPage

	'Edits the information for the selected product.
	
	ElseIf Request("Function") = "4" Then

	MyCurrSymbol = FormatCurrency(1)

	MyCurrSymbol = Replace((MyCurrSymbol),"1","")
	MyCurrSymbol = Replace((MyCurrSymbol),",","")
	MyCurrSymbol = Replace((MyCurrSymbol),".","")
	MyCurrSymbol = Replace((MyCurrSymbol),"0","")
	
	ProdPrice = Replace(Request("Price"),MyCurrSymbol,"")	
	If MyCurrSymbol = "$" Then
	rPrice = Replace((ProdPrice),",","")
	Else
	rPrice = Replace((ProdPrice),",",".")
	End If

	SQLStmt = "UPDATE Product SET Product_ID = '" & Request("PRODUCT_ID") & "', "
	SQLStmt = SQLStmt & "DESCRIPTION = '" & Description & "', "
	SQLStmt = SQLStmt & "PRICE = '" & rPrice & "', "
	SQLStmt = SQLStmt & "MESSAGE = '" & Message & "', "
	SQLStmt = SQLStmt & "CATEGORY = '" & Request("Category") & "', "
	SQLStmt = SQLStmt & "WEIGHT = " & Weight & ", "
	SQLStmt = SQLStmt & "IMAGE_PATH = '" & Request("IMAGE_PATH") & "', "
	SQLStmt = SQLStmt & "LINK = '" & Request("LINK") & "', "
	SQLStmt = SQLStmt & "MFG = '" & Manufacturer & "' "
	SQLStmt = SQLStmt & "WHERE PRODUCT_ID = '" & Request("Product_ID") & "' "
	'Response.Write SQLStmt
	Set RS4 = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * From Product WHERE PRODUCT_ID = '" & Request("Product_ID") & "' "

	Set RS4A = Connection.Execute(SQLStmt)
	
	End If

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>StoreFront - Edit Product</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body style="font-family: Verdana">
<% If Request("Function") = "1" Then %><% 
	CurrentRecord = 0

	Do While NOT RS1.EOF
%>

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Description.value.length > 255)
  {
    alert("Please enter at most 255 characters in the \"Description\" field.");
    theForm.Description.focus();
    return (false);
  }

  if (theForm.Message.value.length > 255)
  {
    alert("Please enter at most 255 characters in the \"Message\" field.");
    theForm.Message.focus();
    return (false);
  }

  if (theForm.Price.value == "")
  {
    alert("Please enter a value for the \"Price\" field.");
    theForm.Price.focus();
    return (false);
  }

  if (theForm.Price.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"Price\" field.");
    theForm.Price.focus();
    return (false);
  }

  var checkOK = "0123456789-.$-,";
  var checkStr = theForm.Price.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit and \".$-,\" characters in the \"Price\" field.");
    theForm.Price.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form action="product.asp?Function=4&amp;Product_ID=<%= RS1("PRODUCT_ID") %>" method="post" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
  <table cellpadding="2" cellspacing="2" width="90%" align="center" border="0">
    <tr>
      <td colspan="4" align="center" bgcolor="#A09A8B"><strong>Edit Product Listing <u>Product ID: <%= RS1("Product_ID") %></strong></u> </td>
    </tr>
     
    <tr>
      <td bgcolor="#A09A8B"><strong>Description:</strong></td>
      <td colspan="3"><strong><!--webbot bot="Validation" s-display-name="Description" s-data-type="String" i-maximum-length="255" --><textarea name="Description" rows="2" cols="50"><%= RS1("Description") %></textarea></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Message:</strong></td>
      <td colspan="3"><strong><!--webbot bot="Validation" i-maximum-length="255" --><textarea name="Message" rows="2" cols="50"><%= RS1("Message") %></textarea></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Image Path:</strong></td>
      <td colspan="3"><strong><input type="text" size="50" name="IMAGE_PATH" value="<%= RS1("IMAGE_PATH") %>" maxlength="125"></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Link:</strong></td>
      <td colspan="3"><strong><input type="text" size="50" name="LINK" value="<%= RS1("LINK") %>" maxlength="125"></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Category:</strong></td>
      <td><strong><input type="text" size="30" name="CATEGORY" value="<%= RS1("CATEGORY") %>" maxlength="30"></strong></td>
      </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Manufacturer:</strong></td>
      <td><strong><input type="text" size="30" name="MANUFACTURER" value="<%= RS1("MFG") %>" maxlength="30"></strong></td>
       <td bgcolor="#A09A8B"><div align="left"><p><strong>Weight: </strong></p>
        </div></td>
      <td align="left"><strong>&nbsp; <input type="text" size="5" value="<%= RS1("WEIGHT") %>" name="WEIGHT" value maxlength="10"></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>Price:</strong></td>
      <td><strong><!--webbot bot="Validation" s-display-name="Price" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=".$-," b-value-required="TRUE" i-minimum-length="4" --><input type="text" size="10" name="Price" value="<%= FormatCurrency(RS1("Price")) %>"></strong></td>
     
      <td colspan="2"><input type="submit" name="Submit" value="Submit"><input type="reset" name="Reset" value="Reset"></td>
     </tr>
  </table>
</form>
</body>
</html>
<%
	RS1.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop

%><% ElseIf Request.QueryString("Function") = "4" Then %><% 
	CurrentRecord = 0

	Do While NOT RS4A.EOF
%><html>

<body>

 <table cellpadding="2" cellspacing="2" width="90%" align="center" border="0">
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="center"><strong>Confirm Product Update</strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>PRODUCT ID</strong></td>
    <td width="30%">&nbsp;<strong><%= RS4A("Product_ID") %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>PRICE</strong></td>
    <td width="30%">&nbsp;<strong><%= FormatCurrency(RS4A("Price")) %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>CATEGORY</strong></td>
    <td width="30%">&nbsp;<strong><% If RS4A("CATEGORY") = "" Then %>None Specified <% Else %>
    <%= RS4A("Category") %><% End If %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>WEIGHT</strong></td>
    <td width="30%">&nbsp;<strong><%= RS4A("Weight") %></strong></td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="left"><strong>DESCRIPTION</strong></td>
  </tr>
  <tr>
    <td colspan="4">&nbsp;<strong><% If RS4A("DESCRIPTION") = "" Then %>None Specified <% Else %>
	<%= RS4A("Description") %><% End If %></strong></td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="left"><strong>MESSAGE</strong></td>
  </tr>
  <tr>
    <td colspan="4">&nbsp;<strong><% If RS4A("MESSAGE") = "" Then %>None Specified <% Else %>
    <%= Server.HTMLEncode(RS4A("MESSAGE")) %><% End If %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>IMAGE PATH</strong></td>
    <td width="30%">&nbsp;<strong><% If RS4A("IMAGE_PATH") = "" Then %>None Specified <% Else %>
    <%= RS4A("IMAGE_PATH") %><% End If %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>LINK</strong></td>
    <td width="30%">&nbsp;<strong><% If RS4A("LINK") = "" Then %>None Specified <% Else %>
    <%= RS4A("LINK") %><% End If %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>MANUFACTURER</strong></td>
    <td width="30%">&nbsp;<strong><% If RS4A("MFG") = "" Then %>None Specified <% Else %>
    <%= RS4A("MFG") %><% End If %></strong></td>
    <td width="20%">&nbsp;</td>
    <td width="30%">&nbsp;</td>
  </tr>
</table>
<%
	RS4A.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop

%><% ElseIf Request.QueryString("Function") = "2" Then %><% 
	CurrentRecord = 0

	Do While NOT RS2A.EOF
%>
<table cellpadding="2" cellspacing="2" width="90%" align="center" border="0">
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="center"><strong>Product Entry Confirmatation</strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>PRODUCT ID</strong></td>
    <td width="30%">&nbsp;<strong><%= RS2A("Product_ID") %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>PRICE</strong></td>
    <td width="30%">&nbsp;<strong><%= FormatCurrency(RS2A("Price")) %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>CATEGORY</strong></td>
    <td width="30%">&nbsp;<strong><% If RS2A("CATEGORY") = "" Then %>None Specified <% Else %>
    <%= RS2A("Category") %><% End If %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>WEIGHT</strong></td>
    <td width="30%">&nbsp;<strong><% If RS2A("WEIGHT") = "" Then %>None Specified <% Else %>
   <%= RS2A("WEIGHT") %><% End If %></strong></td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="left"><strong>DESCRIPTION</strong></td>
  </tr>
  <tr>
    <td colspan="4">&nbsp;<strong><% If RS2A("DESCRIPTION") = "" Then %>None Specified <% Else %>
    <%= RS2A("DESCRIPTION") %><% End If %></strong></td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#A09A8B" align="left"><strong>MESSAGE</strong></td>
  </tr>
  <tr>
    <td colspan="4">&nbsp;<strong><% If RS2A("MESSAGE") = "" Then %>None Specified <% Else %>
    <%= Server.HTMLEncode(RS2A("MESSAGE")) %><% End If %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>IMAGE PATH</strong></td>
    <td width="30%">&nbsp;<strong><% If RS2A("IMAGE_PATH") = "" Then %>None Specified <% Else %>
    <%= RS2A("IMAGE_PATH") %><% End If %></strong></td>
    <td width="20%" bgcolor="#A09A8B"><strong>LINK</strong></td>
    <td width="30%">&nbsp;<strong><% If RS2A("LINK") = "" Then %>None Specified <% Else %>
    <%= RS2A("LINK") %><% End If %></strong></td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#A09A8B"><strong>MANUFACTURER</strong></td>
    <td width="30%">&nbsp;<strong><% If RS2A("MFG") = "" Then %>None Specified <% Else %>
    <%= RS2A("MFG") %><% End If %></strong></td>
    <td width="20%">&nbsp;</td>
    <td width="30%">&nbsp;</td>
  </tr>
</table>
<%
	RS2A.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop

%><% ElseIf Request("Function") = "3" Then %>
<table cellpadding="2" cellspacing="2" width="90%" align="center" border="0">
  <tr>
    <td colspan="2" bgcolor="#A09A8B" align="center"><strong>Product Delete Confirmatation</strong></td>
  </tr>
  <tr>
    <td align="left">&nbsp; Product ID Number: <b><%= Request("Product_ID") %></b> has been deleted</td>
  </tr>
</table>
<% End If %>
<p align="center"><small><a href="prodadd.htm">Add Product</a> | <a href="proddelete.htm">Delete Product</a> | <a href="prodlist.asp">List Products</a> | <a href="prodedit.htm">Edit Product</a><br>
<a href="reports.htm">Sales Reporting</a> | <a href="set_up.asp?Update=0">Store Set-Up</a></small></p>
</body>
</html>









































