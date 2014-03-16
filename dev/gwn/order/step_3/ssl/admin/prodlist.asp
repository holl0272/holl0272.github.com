<%@ LANGUAGE="VBSCRIPT" %>

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.2.5.1
'
'   Author     :   LaGarde, Incorporated

'   Description:   Produces a list of all products contained in the 
'                  products table.
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
%>

<%

Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4


	Dim DSN_Name
	DSN_Name = Session("DSN_Name")

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set RS = Server.CreateObject("ADODB.RecordSet")
	
	Connection.Open "DSN="&DSN_Name&""
	
	SQLStmt = "SELECT * FROM Product"


	RS.Open SQLStmt, Connection, adOpenKeyset,adLockReadOnly

	RS.PageSize = 5

ScrollAction = Request("ScrollAction")
if ScrollAction <> "" Then
	PageNo = mid(ScrollAction, 5)
	if PageNo < 1 Then 
		PageNo = 1
	end if
else
	PageNo = 1
end if
RS.AbsolutePage = PageNo



%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>List All Products</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">


<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none">
</head>

<body>
<div align="center"><center>

  <% 
	RowCount = rs.PageSize
	Do While Not RS.EOF and rowcount > 0 
%>
	
<table cellpadding="2" cellspacing="2" width="90%" align="center" border="0">
    <tr>
      <td width="25%" bgcolor="#A09A8B"><strong>PRODUCT ID</strong></td>
	  <td width="25%" colspan="2">&nbsp;<strong><%= RS("Product_ID") %></strong></td>
	  <td width="25%" bgcolor="#A09A8B"><strong>PRICE</strong></td>
      <td width="25%">&nbsp;<strong><%= FormatCurrency(RS("Price"),2) %></strong></td>
	
    </tr>
    <tr>
	  <td bgcolor="#A09A8B"><strong>CATEGORY</strong></td>
	  <td colspan="2">&nbsp;<strong><% If RS("CATEGORY") = "" Then %> None Specified <% Else %>
          <%=  RS("Category") %><% End If %></strong></td>
	  <td bgcolor="#A09A8B"><strong>WEIGHT</strong></td>
	  <td>&nbsp;<strong><% If RS("Weight") = "" Then %> None Specified <% Else %>
          <%= RS("Weight") %><% End If %></strong></td>
	
    </tr>

    <tr>
      <td colspan="5" align="left" bgcolor="#A09A8B"><strong>DESCRIPTION</strong></td>
    </tr>
    <tr>
      <td colspan="5">&nbsp;<strong><% If RS("Description") = "" Then %> None Specified <% Else %>
      <%= RS("Description") %><% End If %></strong></td>
    </tr>
    <tr>
      <td colspan="5" align="left" bgcolor="#A09A8B"><strong>MESSAGE</strong></td>
    </tr>
    <tr>
      <td colspan="5">&nbsp;<strong><% If RS("Message") = "" Then %> None Specified <% Else %>
      <%= RS("Message") %><% End If %></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>LINK</strong></td>
      <td colspan="4">&nbsp;<strong><% If RS("Link") = "" Then %> None Specified <% Else %>
      <strong><%=  RS("Link") %><% End If %></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>IMAGE PATH</strong></td> 
      <td colspan="4">&nbsp;<strong><% If RS("Image_Path") = "" Then %> None Specified <% Else %>
      <%= RS("Image_Path") %><% End If %></strong></td>
    </tr>
    <tr>
      <td bgcolor="#A09A8B"><strong>MANUFACTURER</strong></td>
      <td colspan="4">&nbsp;<strong><% If RS("MFG") = "" Then %> None Specified <% Else %>
      <%= RS("MFG") %><% End If %></strong></td>
    </tr>
    <tr>
	  <td colspan="3">&nbsp;</td>
	  <td align="center"><form action="product.asp?Function=3&amp;Product_ID=<%= RS("Product_ID") %>" method="Post"><input type="submit" value="DELETE" name="DELETE"></form></td>
      <td align="center"><form action="product.asp?Function=1&amp;Product_ID=<%= RS("Product_ID") %>" method="post"><input type="submit" value="EDIT" name="EDIT"></form></td>
    </tr>
  </table>

  <hr size="2" width="90%" color="000000">
  <%
	RowCount = RowCount - 1
	RS.MoveNext
	Loop
%>
	

  <form METHOD="GET" ACTION="prodlist.asp?">
	
    <% 
	If RS.EOF Then 
	RSEnd = "T"
	ELSE
	RSEnd = "F"
	End If
%>

    <%
'Set RS = RS.NextRecordSet
'Loop

set Connection = nothing
%>



    <% If RSEnd = "F" Then %>
    <% if PageNo > 1 Then %>
    <input TYPE="SUBMIT" NAME="ScrollAction" VALUE="<%="Page " & PageNo-1%>">
    <% end if %>
    <% if RowCount = 0 Then %>
    <input TYPE="SUBMIT" NAME="ScrollAction" VALUE="<%="Page " & PageNo+1%>">
    <% end if %>

    <% Else %>
    <% End If %>


  </form>
  </center></div>
<p align="center"><a href="prodadd.htm">Add Product</a> | <a href="proddelete.htm">Delete Product</a> | <a href="prodlist.asp">List Products</a> | <a href="prodedit.htm">Edit Product</a><br>
<a href="reports.htm">Sales Reporting</a> | <a href="set_up.asp?Update=0">Store Set-Up</a></p>

</body>
</html>
