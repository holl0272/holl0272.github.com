<%
Option Explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: processor.asp
	 

'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	
dim sEndDate,weekStart
If Request.Form("btnSubmit.x") <> "" Then
	If Request.Form("selReport") = "SaleDetails" Then
		If Request.Form("startDate") <> "" and Request.Form("endDate") <> "" Then
			
			sEndDate = Request.Form("endDate")
			sEndDate = dateadd("d",1,sEndDate)
			Response.Redirect "sfReports1.asp?startDate=" & Request.Form("startDate") & "&endDate=" & sEndDate
		End If
	ElseIf Request.Form("selReport") = "SaleSummary" Then
		If Request.Form("startDate") <> "" and Request.Form("endDate") <> "" Then
			sEndDate = Request.Form("endDate")
			sEndDate = dateadd("d",1,sEndDate)
			Response.Redirect "sfReports2.asp?startDate=" & Request.Form("startDate") & "&endDate=" & sEndDate
		End If	
	ElseIf Request.Form("selReport") = "TransactionServiceReport" Then
		If Request.Form("startDate") <> "" and Request.Form("endDate") <> "" Then
			sEndDate = Request.Form("endDate")
			sEndDate = dateadd("d",1,sEndDate)
			Response.Redirect "sfReports3.asp?startDate=" & Request.Form("startDate") & "&endDate=" & sEndDate
		End If	
	ElseIf Request.Form("selReport") = "TradingPartnerReports" Then
		If Request.Form("startDate") <> "" and Request.Form("endDate") <> "" Then
			sEndDate = Request.Form("endDate")
			sEndDate = dateadd("d",1,sEndDate)
			Response.Redirect "sfReports4.asp?startDate=" & Request.Form("startDate") & "&endDate=" & sEndDate
		End If	
	ElseIf Request.Form("selReport") = "ProductSaleReport" Then
		If Request.Form("startDate") <> "" and Request.Form("endDate") <> "" Then
			sEndDate = Request.Form("endDate")
			sEndDate = dateadd("d",1,sEndDate)
			Response.Redirect "sfReports5.asp?startDate=" & Request.Form("startDate") & "&endDate=" & Request.Form("endDate")
		End If	
	ElseIf Request.Form("selReport") = "RetriveOrder" Then
     
		Response.Redirect "sfReports6.asp"
	End If
End If
If DatePart("w", date(), 1) = 1 Then
	weekStart = date()-6
ElseIf DatePart("w", date(), 1) = 2 Then
	weekStart = date()
ElseIf DatePart("w", date(), 1) = 3 Then
	weekStart = date()-1
ElseIf DatePart("w", date(), 1) = 4 Then
	weekStart = date()-2
ElseIf DatePart("w", date(), 1) = 5 Then
	weekStart = date()-3
ElseIf DatePart("w", date(), 1) = 6 Then
	weekStart = date()-4
ElseIf DatePart("w", date(), 1) = 7 Then
	weekStart = date()-5
End If 
%>
<script language="javascript">
function linkChange(start, end){
	var e
	for (i=0;i<document.links.length;i++) {
		e = document.links[i].href
		if (e.indexOf("sfReports")!=-1) {
			if (document.frmReports.selReport.options[0].selected) {
				document.links[i].href = "sfReports1.asp?startDate=" + start + "&endDate=" + end
			}
			if (document.frmReports.selReport.options[1].selected) {
				document.links[i].href = "sfReports2.asp?startDate=" + start + "&endDate=" + end
			}
			if (document.frmReports.selReport.options[2].selected) {
				document.links[i].href = "sfReports3.asp?startDate=" + start + "&endDate=" + end
			}
			if (document.frmReports.selReport.options[3].selected) {
				document.links[i].href = "sfReports4.asp?startDate=" + start + "&endDate=" + end
			}
			if (document.frmReports.selReport.options[4].selected) {
				document.links[i].href = "sfReports5.asp?startDate=" + start + "&endDate=" + end
			}
			if (document.frmReports.selReport.options[5].selected) {
				document.links[i].href = "sfReports6.asp"
			}
		}	
	}
}
function checkLogin(sLogin) {
	if (sLogin == "") {
		alert("The admin folder is unsecured, please contact your network administrator to password protect your admin folder")
	} 
}

function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	

</script>
<html>
<head>
<title>SF Reports Page</title>
</head>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>" onload="checkLogin('<%= Request.ServerVariables("AUTH_USER") %>')">

<form method="post" name="frmReports">
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
	<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">Select a report from the drop down box below.  Enter a date in the Start Date and End Date fields and Submit or choose one of the pre-defined Day, Week to Date, Month to Date or Year to Date reports.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a>
	</td>    
    </tr>
    <tr>
    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
        <table border="0" width="100%" cellpadding="4" cellspacing="0">
        <% If Request.Form("btnSubmit.x") <> "" Then %>
        <tr>
        <td width="100%" colspan="4" align="center" height="90" valign="center">
			<table width="60%" cellpadding="1" cellspacing="0" bgcolor="<%= C_BORDERCOLOR1 %>">
			<tr><td width="100%" align="center">
				<table cellpadding="5" cellspacing="0" bgcolor="<%= C_BGCOLOR4 %>" width="100%">
				<tr><td width="100%" bgcolor="<%= C_BGCOLOR4 %>" align="center" background="<%= C_BKGRND5 %>">
				<font face="<%= C_FONTFACE5 %>" color="#992222" size="<%= C_FONTSIZE5+2 %>"><b>Please Enter a Start and End Date
				</font>
				</td>
				</tr>
				</table>
			</td></tr>	
			</table>
        </td>
        </tr>
		<% End If %>
        <tr>
		<td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>" colspan="4"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1 %>">Create Report</font></b></td>        
		
        </tr>
        <tr>
        <td align="center" valign="top" colspan="4" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">[ <a href="sfReports.asp?startDate=<%= date() %>&amp;endDate=<%= date() %>" onClick="javascript:linkChange('<%= date() %>', '<%= dateadd("d",1,date()) %>')">Day</a> | <a href="sfReports.asp?startDate=<%= weekStart %>&amp;endDate=<%= dateadd("d",1,date()) %>" onClick="javascript:linkChange('<%= weekStart %>', '<%= dateadd("d",1,date()) %>')">Week to Date</a> | <a href="sfReports.asp?startDate=<%= month(date()) & "/01/" & year(date()) %>&amp;endDate=<%= dateadd("d",1,date()) %>" onClick="javascript:linkChange('<%= month(date()) & "/01/" & year(date()) %>', '<%= dateadd("d",1,date()) %>')">Month to Date</a> | <a href="sfReports.asp?startDate=<%= "01/01/" & year(date()) %>&amp;endDate=<%= date()%>" onClick="javascript:linkChange('<%= "01/01/" & year(date()) %>', '<%= dateadd("d",1,date()) %>')">Year to Date</a> ]</font></td>
        </tr>        
        <tr>
        <td align="right" valign="top" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Start Date:</font></td>
        <td nowrap><input name="startDate" style="<%= C_FORMDESIGN %>" size="20"></td>
        <td align="right" valign="top" nowrap><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">End Date:</font></td>
        <td nowrap><input name="endDate" style="<%= C_FORMDESIGN %>" size="20"></td>
        </tr>        
        <tr>
        <td width="100%" align="center" valign="top" colspan="4"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Select Report:</font>
			<select size="1" name="selReport" style="<%= C_FORMDESIGN %>">
                <option value="SaleDetails">Sale Details</option>
                <option value="SaleSummary">Sale Summary</option>
                <option value="TransactionServiceReport">Transaction Service Report</option>
                <option value="TradingPartnerReports">Affiliate Partners Report</option>
                <option value="ProductSaleReport">Product Sale Report</option>
                <option value="RetriveOrder">Retrieve Order</option>
            </select>
        </td>
        </tr>
        <tr>
        <td width="100%" align="center" valign="top" colspan="4"><input type="image" name="btnSubmit" border="0" src="../<%= C_BTN18 %>" alt="Submit" width="108" height="21"></td>
        </tr>
        </table>
    </td>
    </tr>
        <tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
        </tr>
</table>
</td>
</tr>
</table>
</form>
</body>
</html>


