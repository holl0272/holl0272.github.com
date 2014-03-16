<% Option Explicit 
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = true
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/ssmodAdminReports.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

	Call WriteHeader("",True)
	Call ShowAdminPage
	If Response.Buffer Then Response.Flush
    Call ReleaseObject(cnn)

'*******************************************************************************************************************************************

Sub ShowAdminPage

Dim pblnShowAddonReferences
Dim i

pblnShowAddonReferences = True
pblnShowAddonReferences = False
%>
  <table border="0" cellpadding="8" cellspacing="0" ID="Table1">
    <tr>
      <th width="100%" colspan="3" align=center>
      <div colspan="2" class="clsCurrentLocation">Administration Menu</div>
      </th>
    </tr>
    <tr>
      <td valign="top">
  <table border="1" cellpadding="8" cellspacing="0">
    <tr>
      <td valign="top">
<%



%>
	<script type="text/javascript">
		<!--
		function dtreeIndex()
		{
			index++;
			return index;
		}
		
		var index = -1;
		
		d = new dTree('d');
		<% createTreeMenu %>

		document.write(d);

		//-->
	</script>

      </td>
    </tr>
  </table>
      </td>
      <td valign="top">
		This page is only to test the menu and verify the login settings. It should be removed from a production site!
		<p>Testing: Follow each link to the left. If you are not prompted for a login there is a problem.</p>
      </td>
    </tr>
  </table>

<script language="javascript">
function hidePageSection(theImage, strSectionName)
{
	if (theImage == null) return false;
	if (theImage.src.indexOf("images/UI_OM_expand.gif") > 0)
	{
		theImage.src = "images/UI_OM_collapse.gif";
		setCookie("Display" + strSectionName, 1);
	}else{
		theImage.src = "images/UI_OM_expand.gif"
		deleteCookie("Display" + strSectionName);
	}
	showHideElement(document.getElementById("div" + strSectionName + "Content"))
}

function setPageDisplaySettings()
{
	var strSectionHeader;
	var arySectionHeaders = new Array("OrderMenu","ProductMenu","ReportMenu","SiteSettingsMenu","ContentMenu","SupportingSettingsMenu","AdministrativeMenu");
	
	for (var i = 0;  i < arySectionHeaders.length;  i++)
	{
		strSectionHeader = arySectionHeaders[i];
		if (getCookie("Display" + strSectionHeader) == 1)
		{
			hidePageSection(document.getElementById("img" + strSectionHeader), strSectionHeader);
		}
	}
	
}

setPageDisplaySettings();

</script>
<!--#include file="adminFooter.asp"-->
</body>

</html>
<% End Sub %>