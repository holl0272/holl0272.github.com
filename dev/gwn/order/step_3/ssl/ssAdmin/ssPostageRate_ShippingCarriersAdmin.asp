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

Response.Buffer = True

Class clsCarrier
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsProducts
Private pblnError

'database variables

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsProducts)
End Sub

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

Public Property Get rsProducts()
    Set rsProducts = prsProducts
End Property

Public Sub OutputMessage()

Dim i
Dim aError

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
        Else
            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
        End If
    Next 'i

End Sub 'OutputMessage

'***********************************************************************************************

Public Function Load()

dim pstrSQL
dim p_strWhere
dim i
dim sql

	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		pstrSQL = "SELECT * FROM ssShippingCarriers ORDER BY ssShippingCarrierName"
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

Public Function Add()

Dim sql

	sql = "Insert Into ssShippingCarriers (ssShippingCarrierName,ssShippingCarrierUserName,ssShippingCarrierPassword,ssShippingCarrierRateURL,ssShippingCarrierTrackingURL,ssShippingCarrierImagePath)" _
		& " Values ('New Carrier Name','','','','','')"		
	Response.Write "sql: " & sql & "<br />"
	Response.Flush
	cnn.Execute sql,,128

End Function	'Add

'***********************************************************************************************

Public Function Delete()

Dim sql
Dim paryDirty
Dim i
Dim ssShippingCarrierID

	paryDirty = Split(Request.Form("dirty"),",")
	
	'Update the methods
	For i = 0 To UBound(paryDirty)

		ssShippingCarrierID = Trim(paryDirty(i))
		sql = "Delete From ssShippingCarriers Where ssShippingCarrierID=" & ssShippingCarrierID
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
		
		sql = "Delete From ssShippingMethods Where ssShippingCarrierID=" & ssShippingCarrierID
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
		
	Next 'i

End Function	'Add

'***********************************************************************************************

Public Function Update()

Dim sql
Dim strErrorMessage
Dim vItem
Dim paryIDs
Dim paryDirty
Dim i

'On Error Resume Next

    pblnError = False

	paryDirty = Split(Request.Form("dirty"),",")
	
	'For Each vItem in Request.Form
	'	debugprint vItem, Request.Form(vItem)
	'Next

	Dim ssShippingCarrierID
	Dim ssShippingCarrierName
	Dim ssShippingCarrierUserName
	Dim ssShippingCarrierPassword
	Dim ssShippingCarrierRateURL
	Dim ssShippingCarrierTrackingURL
	Dim ssShippingCarrierImagePath

	'Update the methods
	For i = 0 To UBound(paryDirty)

		ssShippingCarrierID = Trim(paryDirty(i))
		ssShippingCarrierName = Request.Form("ssShippingCarrierName" & ssShippingCarrierID)
		ssShippingCarrierUserName = Request.Form("ssShippingCarrierUserName" & ssShippingCarrierID)
		ssShippingCarrierPassword = Request.Form("ssShippingCarrierPassword" & ssShippingCarrierID)
		ssShippingCarrierRateURL = Request.Form("ssShippingCarrierRateURL" & ssShippingCarrierID)
		ssShippingCarrierTrackingURL = Request.Form("ssShippingCarrierTrackingURL" & ssShippingCarrierID)
		ssShippingCarrierImagePath = Request.Form("ssShippingCarrierImagePath" & ssShippingCarrierID)

		sql = "Update ssShippingCarriers Set " _
			& makeSQLUpdate("ssShippingCarrierName", ssShippingCarrierName, False, 0) & ", " _
			& makeSQLUpdate("ssShippingCarrierUserName", ssShippingCarrierUserName, False, 0) & ", " _
			& makeSQLUpdate("ssShippingCarrierPassword", ssShippingCarrierPassword, False, 0) & ", " _
			& makeSQLUpdate("ssShippingCarrierRateURL", ssShippingCarrierRateURL, False, 0) & ", " _
			& makeSQLUpdate("ssShippingCarrierTrackingURL", ssShippingCarrierTrackingURL, False, 0) & ", " _
			& makeSQLUpdate("ssShippingCarrierImagePath", ssShippingCarrierImagePath, False, 0) & " " _
			& " Where  ssShippingCarrierID=" & ssShippingCarrierID

			'Response.Write i & ": " & sql & "<br />"
			'Response.Flush
			cnn.Execute sql,,128
	Next 'i
		
    Update = (not pblnError)

End Function    'Update

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

End Class   'clsCarrier

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'******************************************************************************************************************************************************************

'******************************************************************************************************************************************************************

mstrPageTitle = "Shipping Carriers Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsCarrier
Dim mobjrsCarriers

	mAction = LoadRequestValue("Action")

	Set mobjrsCarriers = GetRS("Select ssShippingCarrierID, ssShippingCarrierName from ssShippingCarriers Order By ssShippingCarrierName")
	
    Set mclsCarrier = New clsCarrier
    With mclsCarrier
		If mAction = "Update" Then .Update
		If mAction = "Add" Then .Add
		If mAction = "Delete" Then .Delete
		.Load
	End With
    
	Call WriteHeader("",True)
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

function CheckAll(blnCheck)
{
	var plngCount;
	var i;

	plngCount = document.frmData.dirty.length;
	if (document.frmData.dirty.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
		document.frmData.dirty[i].checked = blnCheck;
		}
	}else{
	document.frmData.dirty.checked = blnCheck;
	}
	
}

function makeDirty(theItem, lngID)
{

	var plngCount;
	var i;

	plngCount = document.frmData.dirty.length;
	if (document.frmData.dirty.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
			if (document.frmData.dirty[i].value == lngID)
			{
			document.frmData.dirty[i].checked = true;
			return true
			}
		}
	}

}

//-->
</SCRIPT>

<CENTER>

<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
</table>

<FORM action="ssPostageRate_ShippingCarriersAdmin.asp" id="frmData" name="frmData" method="post">
<input type="hidden" id="Action" name="Action" value="Update">

<table class="tbl" width="100%" cellpadding="2" cellspacing="0" border="1" id="tblSummary">
  <tr class="tblhdr">
    <th>&nbsp;</th>
    <th colspan="1" align="left">Carrier Name</th>
    <th colspan="1" align="left">Username</th>
    <th colspan="1" align="left">Password</th>
<!--
    <th colspan="1" align="left">Rate URL</th>
    <th colspan="1" align="left">Tracking URL</th>
    <th colspan="1" align="left">Image Path</th>
-->
  </tr>

  <% 
  Dim plngShippingCarrierID
  
  With mclsCarrier.rsProducts 
	Do While Not .EOF
		plngShippingCarrierID = .Fields("ssShippingCarrierID").Value
  %>
  <tr>
    <td align="left"><input type="checkbox" name="dirty" ID="dirty<%= plngShippingCarrierID %>" value="<%= plngShippingCarrierID  %>"><input type="hidden" NAME="ShippingCarrierID" ID="ShippingCarrierID" value="<%= plngShippingCarrierID  %>"></td>
    <td align="left"><input type="text" name="ssShippingCarrierName<%= plngShippingCarrierID %>" id="ssShippingCarrierName<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierName").Value) %>" onchange="makeDirty(this, <%= plngShippingCarrierID %>);" onblur="return isEmpty(this, 'Please enter a carrier name')" size="20" maxlength="65"></td>
    <td align="left"><input type="text" name="ssShippingCarrierUserName<%= plngShippingCarrierID %>" ID="ssShippingCarrierUserName<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierUserName").Value) %>" size="8" maxlength="65"></td>
    <td align="left"><input type="password" name="ssShippingCarrierPassword<%= plngShippingCarrierID %>" ID="ssShippingCarrierPassword<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierPassword").Value) %>" size="8" maxlength="65"></td>
<!--
    <td align="left"><input type="text" name="ssShippingCarrierRateURL<%= plngShippingCarrierID %>" id="ssShippingCarrierRateURL<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierRateURL").Value) %>" size="20" maxlength="255"></td>
    <td align="left"><input type="text" name="ssShippingCarrierTrackingURL<%= plngShippingCarrierID %>" id="ssShippingCarrierTrackingURL<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierTrackingURL").Value) %>" size="20" maxlength="255"></td>
    <td align="left"><input type="text" name="ssShippingCarrierImagePath<%= plngShippingCarrierID %>" id="ssShippingCarrierImagePath<%= plngShippingCarrierID %>" value="<%= Trim(.Fields("ssShippingCarrierImagePath").Value) %>" size="20" maxlength="255"></td>
-->
  </tr>

  <%
	  .MoveNext
	Loop
  End With
  %>
 <tr class="tblhdr">
	<th align="left" colspan="10">
	  &nbsp;&nbsp;<input class="butn" id="btnCheckAll" name="btnCheckAll" type="button" value="Check All" onclick="CheckAll(true);">
	  &nbsp;&nbsp;<input class="butn" id="btnUnCheckAll" name="btnUnCheckAll" type="button" value="Uncheck All" onclick="CheckAll(false);">
	  <INPUT class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/PostageRate/help_PostageRate.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
      <input class="butn" title="Add a new shipping method" id="btnAdd" name="btnAdd" type="button" onclick="this.form.Action.value='Add'; this.form.submit();" value="Add New">
      <input class="butn" title="delete checked carriers" id="btnDelete" name="btnDelete" type="button" onclick="this.form.Action.value='Delete'; this.form.submit();" value="Delete">
      <input class="butn" title="Save changes" id="btnUpdate" name="btnUpdate" type="submit" value="Save">
    </th>
  </TR>
  <tr>
    <td colspan="16" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
          <td align="center" colspan="2"><a href="ssInstallationPrograms/ssPostageRate2_addon_FedEx_Registration.asp">Set-up FedEx</a></td>
        </tr>
        <tr>
          <td align="center"><a href="ssPostageRate_shippingMethodsAdmin.asp">Configure Shipping Methods</a></td>
        </tr>
      </table>
  
    </td>
  </tr>
</TABLE>

</FORM>

</CENTER>
</BODY>
</HTML>
<%

	Set mclsCarrier = Nothing
	Call ReleaseObject(mobjrsCarriers)
	Call ReleaseObject(cnn)

    Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Function WriteCheckboxValue(vntValue)

	If len(Trim(vntValue) & "") > 0 Then
		If cBool(vntValue) Then Response.Write "CHECKED"
	End If

End Function	'WriteCheckboxValue

'************************************************************************************************************************************

%>