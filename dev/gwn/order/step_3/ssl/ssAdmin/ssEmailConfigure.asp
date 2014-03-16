<%Option Explicit
'********************************************************************************
'*   Order Manager for StoreFront 6.0
'*   Gift Wrap editor
'*   File Version:		1.00.0001
'*   Revision Date:		August 18, 2004
'*
'*   1.00.001 (August 18, 2004)
'*   - Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

Response.Buffer = True


'***********************************************************************************************
'***********************************************************************************************

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/ssclsEmail.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
debug.Enabled = True

mstrPageTitle = "Order Manager Sales Receipt/Packing Slip Editor"
Call WriteHeader("",False)

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mstrAction
Dim maryEmailTemplates
Dim mclsEmail
Dim mblnSaveError
Dim mstrBody
Dim mstrEmailFileText
Dim mstrFileName
Dim mstrSubject
Dim mstrTemplateName
Dim mstrSelectedTemplate
Dim i

	mblnSaveError = False
	
    mstrAction = LoadRequestValue("Action")

	mstrTemplateName = LoadRequestValue("TemplateName")
	mstrSubject = LoadRequestValue("Subject")
	mstrBody = LoadRequestValue("Body")
	mstrSelectedTemplate = LoadRequestValue("SelectedTemplate")
	mstrFileName = LoadRequestValue("FileName")
	
	If False Then
		Response.Write "<fieldset><legend>Values</legend>"
		Response.Write "mstrAction: " & mstrAction & "<BR>"
		Response.Write "mstrTemplateName: " & mstrTemplateName & "<BR>"
		Response.Write "mstrSubject: " & mstrSubject & "<BR>"
		Response.Write "mstrBody: " & mstrBody & "<BR>"
		Response.Write "mstrSelectedTemplate: " & mstrSelectedTemplate & "<BR>"
		Response.Write "mstrFileName: " & mstrFileName & "<BR>"
		Response.Write "</fieldset>"
	End If

	Set mclsEmail = New clsEmail
    
    Select Case mstrAction
        Case "Delete"
			If mclsEmail.DeleteEmailTemplate(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate) Then
				mstrMessage = mstrSelectedTemplate & " was successfully deleted."
			Else
				mstrMessage = "<font color=""red"">There was an error deleting " & mstrSelectedTemplate & ".</font>"
			End If
			mstrSelectedTemplate = ""
			Call mclsEmail.LoadEmailTemplates(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate, maryEmailTemplates)
        Case "New"
			Response.Write "<h2>New</h2>"
			If Len(mstrTemplateName) = 0 Then mstrTemplateName = mstrFileName
			ReDim maryEmailTemplates(0)
			maryEmailTemplates(0) = Array(mstrFileName, mstrSubject, mstrBody, mstrTemplateName)
			If mclsEmail.UpdateEmailTemplate(ssAdminPath & cstrEmailTemplateFolder, mstrFileName, 0, maryEmailTemplates) Then
				mstrMessage = mstrTemplateName & " was successfully created as the file " & mstrFileName & "."
			Else
				mstrMessage = "<font color=""red"">There was an error creating " & mstrFileName & ".</font>"
			End If
			mstrSelectedTemplate = mstrFileName
			Call mclsEmail.LoadEmailTemplates(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate, maryEmailTemplates)
        Case "Update"
			Call mclsEmail.LoadEmailTemplates(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate, maryEmailTemplates)
			For i = 0 To UBound(maryEmailTemplates)
				If CBool(maryEmailTemplates(i)(0) = mstrSelectedTemplate) Then
					If Len(mstrTemplateName) = 0 Then mstrTemplateName = mstrFileName
					maryEmailTemplates(i)(0) = mstrFileName
					maryEmailTemplates(i)(1) = mstrSubject
					maryEmailTemplates(i)(2) = mstrBody
					maryEmailTemplates(i)(3) = mstrTemplateName
				
					If mclsEmail.UpdateEmailTemplate(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate, i, maryEmailTemplates) Then
						mstrMessage = mstrSelectedTemplate & " was successfully updated."
					Else
						mstrMessage = "<font color=""red"">There was an error updating " & mstrSelectedTemplate & ".</font>"
					End If
					Exit For
				End If
			Next 'i

            'mblnSaveError = Not mclsQBConfig.Update
        Case Else
			Call mclsEmail.LoadEmailTemplates(ssAdminPath & cstrEmailTemplateFolder, mstrSelectedTemplate, maryEmailTemplates)
    End Select
	
    'Now load the values
	mstrFileName = maryEmailTemplates(0)(0)
	mstrSubject = maryEmailTemplates(0)(1)
	mstrBody = maryEmailTemplates(0)(2)
	mstrTemplateName = maryEmailTemplates(0)(3)
	
    For i = 0 To UBound(maryEmailTemplates)
		If CBool(maryEmailTemplates(i)(0) = mstrSelectedTemplate) Then
			mstrFileName = maryEmailTemplates(i)(0)
			mstrSubject = maryEmailTemplates(i)(1)
			mstrBody = maryEmailTemplates(i)(2)
			mstrTemplateName = maryEmailTemplates(i)(3)
			Exit For
		End If
	Next 'i
    
	mstrEmailFileText = mclsEmail.EmailFileText(mstrTemplateName, mstrSubject, mstrBody)
	Set mclsEmail = Nothing
    
    Call WriteHeader("body_onload();", False)

%>

<script language="javascript" type="text/javascript">
<!--
var blnIsDirty;

function body_onload()
{
	blnIsDirty = false;
}

function changeEmailTemplate(theSelect)
{
var theForm = theSelect.form;

	theForm.Action.value = "";
	theForm.submit();
}

function btnDelete_onclick(theButton)
{

var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete template " + getSelectText(theForm.SelectedTemplate) + "?");
    if (blnConfirm)
    {
		theForm.Action.value = "Delete";
		theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnNew_onclick(theButton)
{

var theForm = theButton.form;
var blnConfirm;

	theForm.Action.value = "New";

	theForm.FileName.disabled = false;
	theForm.FileName.value = "Enter a file name with a .txt extension";
	theForm.FileName.select();
	theForm.FileName.focus();
}

function btnUpdate_onclick(theButton)
{

var theForm = theButton.form;

    if (ValidInput(theForm))
    {
		if(theForm.Action.value == ""){theForm.Action.value = "Update";}
		theForm.submit();
		return(true);
    }
}

function setSelectedEmailReplacement(theSpan)
{
	//alert(theSpan.innerText);
}

function ValidInput(theForm)
{

    return(true);
}
//-->
</script>
<script language="vbscript">

	Function saveFile(byRef strFileName, byRef strContents)
	'Initialize and script ActiveX controls not marked as safe needs to be set to Enable or Promopt

	Dim fso
	Dim MyFile
	Dim pstrErrorMessage
	
		On Error Resume Next
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'trim trailing returns
		'If Right(strContents, 1) = Chr(10) Then strContents = Left(strContents, Len(strContents)-1)
		'If Right(strContents, 1) = Chr(13) Then strContents = Left(strContents, Len(strContents)-1)
		'If Right(strContents, 1) = vbcrlf Then strContents = Left(strContents, Len(strContents)-1)

		If Err.number = 429 Then
			pstrErrorMessage = "You do not have the security settings set properly for this item. " & vbcrlf _
							& "To enable this functionality do the following: "  & vbcrlf _
							& "  - In the Internet Explorer toolbar select Tools --> Internet Options "  & vbcrlf _
							& "  - Select the security tab "  & vbcrlf _
							& "  - Select Custom Level "  & vbcrlf _
							& "  - Find the option 'Initialize and Script ActiveX Components not marked as safe' "  & vbcrlf _
							& "    Change this setting to Prompt "  & vbcrlf _
							& "  - Select OK and OK "
			msgbox(pstrErrorMessage)
		ElseIf Err.number > 0 Then
			pstrErrorMessage = "There was an error opening the file " & pstrFilePath & ". " & vbcrlf _
							& "Error " & Err.number & ": " & Err.Description & vbcrlf
			msgbox(pstrErrorMessage)
		Else
			Set MyFile = fso.CreateTextFile(strFileName, True)
			MyFile.Write strContents
			MyFile.close
			Set MyFile = Nothing
		End If

		Set fso = Nothing
		
		saveFile = (Err.number = 0)

	End Function	'saveFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function saveExportFile(strExportSection)

	Dim pobjFileChooser
	Dim pobjExportSection
	Dim pstrFile
	Dim pstrTextToSave

		Set pobjExportSection = document.all(strExportSection)
		Set pobjFileChooser = document.all("tempFile")
		pobjFileChooser.click()

		pstrFile = pobjFileChooser.value
		If Len(pstrFile) > 0 Then
			pstrTextToSave = pobjExportSection.value
			msgbox pstrTextToSave
			Call saveFile(pstrFile, pstrTextToSave)
		End If

		saveExportFile = False

		Set pobjFileChooser = Nothing
		Set pobjExportSection = Nothing
		
		saveExportFile = (Err.number <> 0)

	End Function	'saveExportFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

</script>

<%= mstrMessage %>
<% If mblnSaveError Then %>
<h4><font color=red>There was an error saving the configuration file. This is likely due to the security settings on your server. You should manually save the file by clicking this link. Note you will need to upload it to the server as you can only save it to your local workstation.</font></h4>
<a href="" onclick="saveExportFile('QuickBooksConfig'); return false;">Manually save ssAdmin\exportTemplates\supportingTemplates\orderManager_TemplateConfiguration.xsl</a><br>
<% End If %>
<FORM action="ssEmailConfigure.asp" id="frmData" name="frmData" onsubmit="return ValidInput();" method="post">
<input type="hidden" name="Action" id="Action" value="" />
<table class="tbl" cellpadding="3" cellspacing="0" border="1" style="border-collapse: collapse" id="tblData">
  <colgroup>
    <col align="right" valign="top" />
    <col align="left" valign="top" />
  </colgroup>

	<tr class="tblhdr">
		<th colspan="2" align="center">Order Manager Email Templates</th>
	</tr>

	<tr>
		<td><label for="SelectedTemplate" title="">Email Template</label>: </td>
		<td>
		<select name="SelectedTemplate" id="SelectedTemplate" onchange="changeEmailTemplate(this);">
		<% For i = 0 To UBound(maryEmailTemplates) %>
		<option value="<%= maryEmailTemplates(i)(0) %>" <%= isSelected(maryEmailTemplates(i)(0) = mstrSelectedTemplate) %>><%= maryEmailTemplates(i)(3) %></option>
		<% Next 'i %>
		</select>&nbsp;&nbsp;&nbsp;<label for="FileName" title="">File&nbsp;Name</label>:&nbsp;<input name="FileName" id="FileName" value="<%= mstrFileName %>" size="45" disabled />
		</td>
  	</tr>
	<tr>
		<td><label for="TemplateName" title="">Template Name</label>: </td>
		<td><input name="TemplateName" id="TemplateName" value="<%= mstrTemplateName %>" size="90" /></td>
  	</tr>
	<tr>
		<td><label for="Subject" title="">Subject</label>: </td>
		<td><input name="Subject" id="Subject" value="<%= mstrSubject %>" size="90" /></td>
  	</tr>
	<tr>
		<td valign="top">
		  <label for="Body" title="">Body</label>:&nbsp;<br />
		  <hr />
		  <fieldset onmouseover="showElement(document.all('divEmailReplacements'));" onmouseout="hideElement(document.all('divEmailReplacements'));"><legend><u>Replacement Codes</u>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager_customEmails.htm')" id="btnHelp" name=btnHelp></legend>
		  <div id="divEmailReplacements" style="display:none;" align="center">
		  <font size="1pt">
			<span title="Order Number" onmousedown="setSelectedEmailReplacement(this);">{orderNumber}</span><br />
			<span title="FirstName + MI +LastName" onmousedown="setSelectedEmailReplacement(this);">{customerName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerFirstName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerMI}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerLastName}</span><br />
			<span title="address (both lines) + city, state ZIP" onmousedown="setSelectedEmailReplacement(this);">{customerAddress}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerAddress1}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerAddress2}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerCity}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerState}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerZip}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerCountry}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerCountryName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerPhone}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerFax}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{customerEmail}</span><br />
			<span title="FirstName + MI +LastName" onmousedown="setSelectedEmailReplacement(this);">{recipientName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientFirstName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientMI}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientLastName}</span><br />
			<span title="address (both lines) + city, state ZIP" onmousedown="setSelectedEmailReplacement(this);">{recipientAddress}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientAddress1}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientAddress2}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientCity}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientState}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientZip}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientCountry}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientCountryName}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientPhone}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientFax}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{recipientEmail}</span><br />
			<span title="SiteURL + OrderDetail.aspx?OrderID= + orderUID" onmousedown="setSelectedEmailReplacement(this);">{orderDetailLink}</span><br />
			<span title="SiteURL + OrderTracking.aspx?OrderID= + orderUID" onmousedown="setSelectedEmailReplacement(this);">{trackingLink}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{dateShipped}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{trackingNumber}</span><br />
			<span title="" onmousedown="setSelectedEmailReplacement(this);">{trackingMessage}</span><br />
		  </font>
		  </div>
		  </fieldset>
		</td>
		<td valign="top">
		  <textarea name="Body" id="Body" rows="20" cols="70"><%= mstrBody %></textarea>&nbsp;
		  <a href="javascript:doNothing()" onClick="return openACE(document.frmData.Body);" title="Edit this field with the HTML Editor"><img src="images/source.gif" border="0" /></a>
		</td>
  	</tr>

  <TR>
    <TD>&nbsp;</TD>
    <TD align="left">&nbsp;&nbsp;
        <INPUT class="butn" name="btnReset" id="btnReset" type="reset" value="Reset">&nbsp;&nbsp;
        <INPUT class="butn" name="btnDelete" id="btnDelete" type="button" value="Delete" onclick="btnDelete_onclick(this);">&nbsp;&nbsp;
        <INPUT class="butn" name="btnNew" id="btnNew" type="button" value="New" onclick="btnNew_onclick(this);">&nbsp;&nbsp;
        <INPUT class="butn" name="btnUpdate" id="btnUpdate" type="button" value="Update" onclick="btnUpdate_onclick(this);">
    </TD>
  </TR>
</TABLE>
</FORM>

<% If mblnSaveError Then %>
<span id=spantempFile style="display:none"><input type=file id=tempFile name=tempFile size="20"></span>
<textarea name="QuickBooksConfig" id="QuickBooksConfig" rows="57" cols="120"><%= mstrEmailFileText %></textarea>
<% End If %>
<!--#include file="adminFooter.asp"-->
</BODY>
</HTML>
