<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: sfreports4.asp
	
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
<HTML>
<HEAD>
<title>SF Reports Page</title>
</HEAD>
<SCRIPT language="javascript">
function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2c.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	
</script>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->

<%
Dim sStartDate, sEndDate, rsSales, sSQL, sTotalNet, sTotalSTax, sTotalCTax, sTotalShipping, sGrandTotal, arrSales, sHandling, i, rsPartners, j, sAffiliate, rsAff, sFilter

sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))
sAffiliate = Request.QueryString("Affiliate")

If sAffiliate = "" Then
	Set rsSales = CreateObject("ADODB.RecordSet")
	sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal, orderTradingPartner FROM sfOrders WHERE (orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & ") AND (orderTradingPartner IS NOT NULL) and orderIsComplete = 1 and orderTradingPartner in (select affName from sfAffiliates)"
	If cblnSQLDatabase Then
		sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal, affName As orderTradingPartner" _
			& " FROM sfOrders INNER JOIN sfAffiliates ON sfOrders.orderTradingPartner = convert(varchar(100), sfAffiliates.affID)" _
			& " WHERE (orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & ") AND (orderTradingPartner IS NOT NULL) and (orderIsComplete = 1)"
	Else
		sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal, affName As orderTradingPartner" _
			& " FROM sfOrders INNER JOIN sfAffiliates ON sfOrders.orderTradingPartner = CStr(sfAffiliates.affID)" _
			& " WHERE (orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & ") AND (orderTradingPartner IS NOT NULL) and (orderIsComplete = 1)"
	End If
	rsSales.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText

	If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 

	closeObj(rsSales)
	
	sSQL = "SELECT DISTINCT orderTradingPartner FROM sfOrders WHERE orderTradingPartner <> '' and orderIsComplete = 1 and orderTradingPartner in (select affName from sfAffiliates)"
	Set rsPartners = CreateObject("ADODB.Recordset")
	rsPartners.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
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
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Affiliate Sales Summary</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">Referral sales for all affiliate partners are listed below.  Chose an <B>Affiliate ID</B> to view that affiliate's transaction detail.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        
	        <% 
	        If isArray(arrSales) AND Not (rsPartners.EOF And rsPartners.BOF) Then 
	        %>
	        <tr>
			<td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>" colspan="4"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Report for <%= sStartDate %> to <%= sEndDate %></font></b></td>        
	        </tr>
	   
	        <%
				Do While Not rsPartners.EOF
						sTotalNet = 0 
						sTotalSTax = 0
						sTotalCTax = 0
						sTotalShipping = 0
						sHandling = 0 
						sGrandTotal = 0

						For i=0 to uBound(arrSales, 2)
							If rsPartners.Fields("orderTradingPartner") = arrSales(6, i) Then			
								If arrSales(0, i) <> "" Then sTotalNet = sTotalNet + cDbl(arrSales(0, i))
								If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
								If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
								If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
								If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
								If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
							End If
						Next
	        %>
	        <tr>
	        <td>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
	        <td width="100%" colspan=2 align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b>Affiliate Partner:&nbsp;<A href="sfReports4.asp?Affiliate=<%= rsPartners.Fields("orderTradingPartner") %>"><%= rsPartners.Fields("orderTradingPartner") %></a></b></font></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Net Sales:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sTotalNet) %></font></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total State/Providence Tax:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sTotalSTax) %></font></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total Country Tax:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sTotalCTax) %></font></td>
	        </tr>  
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total Shipping:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sTotalShipping) %></font></td>
	        </tr>                 
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total Handling:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sHandling) %></font></td>
	        </tr>   
	        <tr>
	        <td width="75%" align="right" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Total Sales:</font></td>
	        <td width="25%" align="left" valign="top"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= FormatCurrency(sGrandTotal) %></font></td>
	        </tr>
	        </table>
	        </td>
	        </tr>
	        <%
					rsPartners.MoveNext 
				Loop 
				closeObj(rsPartners)
	        Else 
	        %>
							<tr>
				<td colspan=4 align="center" bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>"><font face="<%= C_FONTFACE5 %>" color="#ff0000" size="<%= C_FONTSIZE5+1 %>">There Were No Affiliate Sales Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>

			<% End If %>
			<tr>
			<td width="100%" align="center" valign="top" colspan="4"></td>
			</tr>
			</table>
<%
Else
	Set rsAff = CreateObject("ADODB.Recordset")
	rsAff.Open "sfAffiliates", cnn, adOpenStatic, adLockReadOnly, adCmdTable
	sFilter = "affName = '" & sAffiliate & "'"
	rsAff.Filter = sFilter	
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
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Affiliate Information</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>"><b>Instructions: </b>The information for the affiliate partner you selected is listed below.  To modify this information, return to the main menu and click on <B>Affiliate Partner Administration</b>.</font></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	    <table width="100%">
	    <%If Not (rsAff.EOF And rsAff.BOF) Then%>
	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Affiliate ID:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affName") %></b></font></td>
	    </tr>
	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Company:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affCompany") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Address Line 1:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affAddress1") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Address Line 2:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affAddress2") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">City:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affCity") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">State:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affState") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Zip/Postal Code:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affZip") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Country:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affCountry") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Phone Number:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affPhone") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Fax Number:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affFAX") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Email:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><a href="mailto:<%= rsAff.Fields("affEmail") %>"><%= rsAff.Fields("affEmail") %></a></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Notes:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b><%= rsAff.Fields("affNotes") %></b></font></td>
	    </tr>
	    	    <tr>
	    <td align="left" width="20%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Web Site:</font></td>
	    <td align="left" width="80%" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><a href="<%= rsAff.Fields("affHttpAddr") %>"><%= rsAff.Fields("affHttpAddr") %></a></font></td>
	    </tr>
	    <% Else %>
	    <tr>
	    <td align="center" width=100%"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4+1 %>">There is no information available for <%= sAffiliate %></font></td>
	    </tr>	    
	    <% End If %>
	    </table>
<%
	closeObj(rsAff)
End If
Call cleanup_dbconnopen	'This line needs to be included to close database connection
%>
    </td>
    </tr>
         <tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
</table>
</td>
</tr>
</table>

</BODY>
</HTML>










