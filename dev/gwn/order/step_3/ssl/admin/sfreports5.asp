<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: sfreports5.asp
	 
'Access Version
'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	
%>
<html>
<head>
<title>SF Reports Page</title>
</head>
<SCRIPT language="javascript">
function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2e.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	
</script>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->

<!--#include file="incAdmin.asp"-->

<%
If Request.Form("btnSubmit.x") <> "" Then
	Dim sStartDate, sEndDate, rsSales, sSQL, sTotalNet, sTotalSTax, sTotalCTax, sTotalShipping, sGrandTotal, arrSales, sHandling, i, sProdId
    Dim Itot 
    Dim iTotalSold 'As Integer
	sStartDate = MakeUSDate(Request.Form("startDate"))
	sEndDate = MakeUSDate(Request.Form("endDate"))
	sProdId = Request.Form("txtProdID")
	If sProdId = "" Then sProdId = Request.Form("sltProdId")
	
	Set rsSales = CreateObject("ADODB.RecordSet")
	 sSql = "SELECT sfProducts.prodName, sfProducts.prodID, sfProducts.prodPrice,sfOrderDetails.odrdtQuantity FROM" _
		  & " (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
		  & " WHERE ((orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & ") AND odrdtProductID = '" & sProdId & "') and sfOrders.orderIsComplete = 1"

'	sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal FROM sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId " _
'	 & "WHERE ((orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & ") AND odrdtProductID = '" & sProdId & "') and sfOrders.orderIsComplete = 1"
	rsSales.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText

	If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 

	sTotalNet = 0 
	'sTotalSTax = 0
	'sTotalCTax = 0
	'sTotalShipping = 0
	'sHandling = 0 
	'sGrandTotal = 0

	rsSales.Close 
	Set rsSales= Nothing
  ITot = 0
	If isArray(arrSales) Then
		For i=0 to uBound(arrSales, 2)
	If arrSales(3, i) <> "" Then 
			 'sTotalNet = sTotalNet + cDbl(arrSales(2, i))
			 iTotalSold = iTotalSold + CInt(arrSales(3, I))
			end if
			'If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
			'If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
			'If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
			'If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
			'If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
		Next
		sTotalNet = CDbl(arrSales(2, 0)) * iTotalSold
		ITot = iTotalSold
	End If 

	%>
	<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
	<table border="0" cellpadding="1" cellspacing="0" bgcolor="<%= C_BORDERCOLOR1 %>" width="<%= C_WIDTH %>" align="center">
	<tr>
	<td>
	    <table width="100%" border="0" cellspacing="1" cellpadding="3">
	    <tr>
	<%	If C_BNRBKGRND = "" Then %>
			<td align="middle" background="<%= C_BKGRND1 %>" bgcolor="<%= C_BGCOLOR1 %>"><b><font face="<%= C_FONTFACE1 %>" color="<%= C_FONTCOLOR1 %>" SIZE="<%= C_FONTSIZE1 %>"><%= C_STORENAME %></font></b></td>
	<%	Else %>
			<td align="middle" bgcolor="<%= C_BNRBGCOLOR %>"><img src="<%= C_BNRBKGRND %>" border="0"></td>
	<%	End If %>        
	    </tr>
	    <tr>
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Sales Summary</font></b></td>        
	    </tr>

	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">Total sales for the product item selected are shown below.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
			<td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>" colspan="4"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Report for Product <%= sProdId %> from <%= sStartDate %> to <%= sEndDate %></font></b></td>        
	        </tr>
	        <% If isArray(arrSales) Then %>
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total Sold:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= iTot %></font></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Net Sales:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sTotalNet) %></font></td>
	        </tr>
	        <% Else %>
	        <tr>
	        <td colspan="2" align="center"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4+1 %>">No Sales Reported for Product <%= sProdId%> from <%= sStartDate %> to <%= sEndDate %></font></td>
	        </tr>
	        <% End If %>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="4"></td>
	        </tr>
	        </table>
	    </td>
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>

	</table>
	</td>
	</tr>
	</table>

	</body>
<% 
Else 
	Dim objRS, sProdList
	Set objRS = getProductList()
	Do While Not objRS.EOF 		
		sProdList = sProdList & "<option value=""" & objRS("prodID") & """>" &  objRs("prodName") & "</option>"
		objRS.MoveNext
	Loop
	closeobj(objRS)
%>
	<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

	<form method="post" name="frmProductSummary">
	<table border="0" cellpadding="1" cellspacing="0" bgcolor="<%= C_BORDERCOLOR1 %>" width="<%= C_WIDTH %>" align="center">
	<tr>
	<td>
	    <table width="100%" border="0" cellspacing="1" cellpadding="3">
	    <tr>
	<%	If C_BNRBKGRND = "" Then %>
			<td align="middle" background="<%= C_BKGRND1 %>" bgcolor="<%= C_BGCOLOR1 %>"><b><font face="<%= C_FONTFACE1 %>" color="<%= C_FONTCOLOR1 %>" SIZE="<%= C_FONTSIZE1 %>"><%= C_STORENAME %></font></b></td>
	<%	Else %>
			<td align="middle" bgcolor="<%= C_BNRBGCOLOR %>"><img src="<%= C_BNRBKGRND %>" border="0"></td>
	<%	End If %>        
	    </tr>
	    <tr>
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">StoreFront Reports</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">Enter the Product ID of the item in the <B>Product ID</B> field or select a product item from the <B>Product</B> drop-down box to view the total sales for the selected item within the date range specified.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
			<td colspan="2" width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>" colspan="4"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1 %>">Create Product Report</font></b></td>        
	        </tr>
            <tr>
            <td width="50%" align="right"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Product ID:</font></td>
            <td width="50%"><input name="txtProdID" style="<%= C_FORMDESIGN %>" size="25"></td>
            </tr>
            <tr>
            <td width="50%" align="right"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Product :</font></td>
            <td width="50%"><select name="sltProdID" style="<%= C_FORMDESIGN %>" size="1"><option></option><%= sProdList %></select></td>
            </tr>
	        <tr>
	        <td colspan="2" width="100%" align="center" valign="top" colspan="4"><input type="image" name="btnSubmit" border="0" src="../<%= C_BTN18 %>" alt="Submit" WIDTH="108" HEIGHT="21"></td>
	        </tr>
	        </table>
	    </td>
	    </tr>
		    <tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>

	</table>
	</td>
	</tr>
	</table>
	<input Type="hidden" name="startDate" value="<%= Request.QueryString("startDate") %>">
	<input Type="hidden" name="endDate" value="<%= Request.QueryString("endDate") %>">
	</form>
<% End If %>	
</html>







