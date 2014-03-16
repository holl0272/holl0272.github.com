<%Option Explicit
'********************************************************************************
'*   Postage Rate Administration						                        *
'*   Release Version: 2.0			                                            *
'*   Release Date: September 21, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'***********************************************************************************************

Function Load()

dim pstrSQL
dim p_strWhere
dim i
dim sql

	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		pstrSQL = "SELECT sfProducts.prodID, sfCategories.catName, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodEnabledIsActive, sfProducts.version, sfProducts.releaseDate, sfProducts.prodFileName" _
				& " FROM sfCategories INNER JOIN sfProducts ON sfCategories.catID = sfProducts.prodCategoryId" _
				& " ORDER BY sfProducts.prodEnabledIsActive Desc, sfProducts.prodName, sfCategories.catName"

		'debugprint "pstrSQL",pstrSQL
		'Response.Flush	  
		  
		On Error Resume Next
		If Err.number <> 0 Then Err.Clear
		
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			Response.Write "<h3><font color=red>The Postage Rate add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
			Response.Write "<a href='ssInstallationPrograms/ssPostageRate2_addon_DBUpgradeTool.asp'>Click here to upgrade</a></h3>"
			Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
			Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
			Response.Flush
			Err.Clear
			Load = False
			Exit Function
		End If
		On Error Goto 0
		
	End With

    Load = (Not prsProducts.EOF)

End Function    'Load

'***********************************************************************************************

Function ConvertBoolean(vntValue)

	If Len(Trim(vntValue & "")) = 0 Then
		ConvertBoolean = False
	Else
		On Error Resume Next
		ConvertBoolean = cBool(vntValue)
		If Err.number <> 0 Then 
			ConvertBoolean = False
			Err.Clear
		End If
	End If

End Function	'ConvertBoolean

'******************************************************************************************************************************************************************

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'******************************************************************************************************************************************************************

'******************************************************************************************************************************************************************

mstrPageTitle = "Product File Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim prsProducts

	mAction = LoadRequestValue("Action")

    Call Load
    
	Call WriteHeader("",True)
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<CENTER>

<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
</table>

<table class="tbl" width="100%" cellpadding="2" cellspacing="0" border="1" id="tblSummary">
  <tr class="tblhdr">
    <th colspan="1" align="left">ID</th>
    <th colspan="1" align="left">Product</th>
    <th colspan="1" align="left">Category</th>
    <th colspan="1" align="left">Price</th>
    <th colspan="1" align="left">Enabled</th>
    <th colspan="1" align="left">Version</th>
    <th colspan="1" align="left">Release Date</th>
    <th colspan="1" align="left">File Path</th>
  </tr>

  <% 
  Dim pobjFSO
  Dim cstrBaseFilePath
  Dim pstrFilePath
  
	If instr(1,Request.ServerVariables("HTTP_HOST"),"localhost") > 0 then
		cstrBaseFilePath = "D:\Sandshot Software\WebSite-New\sandshot\home\InetPub\data\programs\"
	Else
		cstrBaseFilePath = "D:\webppliance\conf\domains\crudmaster\Inetpub\data\programs\"
		cstrBaseFilePath = "D:\users\crudmaster\Inetpub\data\programs\"
	End If
	debugprint "APPL_PHYSICAL_PATH", request.ServerVariables("APPL_PHYSICAL_PATH")
	Set pobjFSO = server.CreateObject("Scripting.FileSystemObject")

  
  With prsProducts
	Do While Not .EOF
 		pstrFilePath = Trim(cstrBaseFilePath & .Fields("prodFileName").Value)
 %>
  <tr>
    <td colspan="1" align="left"><%= .Fields("prodID").Value %></td>
    <td colspan="1" align="left"><%= .Fields("prodName").Value %></td>
    <td colspan="1" align="left"><%= .Fields("catName").Value %>&nbsp;</td>
    <td colspan="1" align="left"><%= .Fields("prodPrice").Value %>&nbsp;</td>
    <td colspan="1" align="left"><%= .Fields("prodEnabledIsActive").Value %>&nbsp;</td>
    <td colspan="1" align="left"><%= .Fields("version").Value %>&nbsp;</td>
    <td colspan="1" align="left"><%= .Fields("releaseDate").Value %>&nbsp;</td>
    <% If Len(.Fields("prodFileName").Value & "") = 0 Then %>
		<td colspan="1" align="left" bgcolor="yellow">-</td>
    <% ElseIf pobjFSO.FileExists(pstrFilePath) Then %>
		<td colspan="1" align="left"><%= .Fields("prodFileName").Value %>&nbsp;</td>
    <% Else %>
		<td colspan="1" align="left" bgcolor="red"><%= .Fields("prodFileName").Value %>&nbsp;</td>
    <% End If %>
  </tr>

  <%
	  .MoveNext
	Loop
  End With
  %>
</TABLE>

</FORM>

</CENTER>
</BODY>
</HTML>
<%

	Call ReleaseObject(prsProducts)
	Call ReleaseObject(cnn)

    Response.Flush

%>