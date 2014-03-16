<% Option Explicit
'********************************************************************************
'*   Gift Certificate Manager				                                    *
'*   Release Version:   1.01.001												*
'*   Release Date:		November 15, 2002										*
'*   Revision Date:		December 5, 2003										*
'*                                                                              *
'*   Release Notes: See ssGiftCertificate_class.asp                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

'***********************************************************************************************

Function SummaryFilter

Dim pstrsqlWhere
Dim pstrsqlHaving
Dim pstrOrderBy

	pstrsqlHaving = " HAVING (ssGCCode<>'')"

	'load the text filter
	mbytText_Filter = Request.Form("optText_Filter")
	mstrText_Filter = Request.Form("Text_Filter")
	If len(mstrText_Filter) > 0 Then
		Select Case mbytText_Filter
			Case "0"	'Do Not Include
			Case "1"	'Certificate #
				pstrsqlHaving = pstrsqlHaving & " AND (ssGCCode Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "2"	'Issued To
				pstrsqlHaving = pstrsqlHaving & " AND (sfCustomers.custLastName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "3"	'Sending Order #
				pstrsqlWhere = " WHERE (CStr(ssGCRedemptionOrderID & '') Like '%" & mstrText_Filter & "%')"
			Case "4"	'Sending email
				pstrsqlHaving = pstrsqlHaving & " AND (ssGCIssuedToEmail Like '%" & mstrText_Filter & "%')"
		End Select	
	End If
	'pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCCustomerID, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _

	mstrStartDate = Request.Form("StartDate")
	If len(mstrStartDate) > 0 then pstrsqlHaving = pstrsqlHaving & " and (ssGCCreatedOn >= " & sqlDateWrap(mstrStartDate & " 12:00:00 AM") & ")"
	
	mstrEndDate = Request.Form("EndDate")
	If len(mstrEndDate) > 0 then pstrsqlHaving = pstrsqlHaving & " and (ssGCCreatedOn <= " & sqlDateWrap(mstrEndDate & " 11:59:59 PM") & ")"

	'load the radio filters
	mbytoptFlag_Active	= Request.Form("optFlag_Active")

	'load the price filter
	mcurLowerPrice = Request.Form("LowerPrice")
	If Len(mcurLowerPrice) = 0 Then
		If Len(Request.Form) = 0 Then mcurLowerPrice = 0.01
	End If
	If Not isNumeric(mcurLowerPrice) Then mcurLowerPrice = ""
	mcurUpperPrice = Request.Form("UpperPrice")
	If Not isNumeric(mcurUpperPrice) Then mcurUpperPrice = ""
	If len(mcurLowerPrice) > 0 then pstrsqlHaving = pstrsqlHaving & " and Sum(ssGCRedemptionAmount)>=" & mcurLowerPrice
	If len(mcurUpperPrice) > 0 then pstrsqlHaving = pstrsqlHaving & " and Sum(ssGCRedemptionAmount)<=" & mcurUpperPrice

	Select Case mbytoptFlag_Active
		Case "0"	'Do Not Include
		Case "1"	'Active
			If cblnSQLDatabase Then
				pstrsqlHaving = pstrsqlHaving & " and (ssGCRedemptionActive=1)"
			Else
				pstrsqlHaving = pstrsqlHaving & " and (ssGCRedemptionActive=1 OR ssGCRedemptionActive=-1)"
			End If
		Case "2"	'Inactive
			pstrsqlHaving = pstrsqlHaving & " and (ssGCRedemptionActive=0 OR ssGCRedemptionActive is Null)"
	End Select	

	'Build  the Order By
	mstrOrderBy = Request.Form("OrderBy")
	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	
	mstrSortOrder = Request.Form("SortOrder")
	If len(mstrSortOrder) = 0 Then mstrSortOrder = "Desc"

	dim paryOrderBy(4)
	paryOrderBy(1) = "ssGCCode"	'Default
	paryOrderBy(2) = "ssGCIssuedToEmail"
	paryOrderBy(3) = "Sum(ssGCRedemptionAmount)"
	paryOrderBy(4) = "ssGCCreatedOn"

	'pstrsqlWhere = " HAVING (ssGCCode<>'') AND (Sum(ssGCRedemptionAmount)>10) AND (ssGCRedemptionActive=0)"
	'pstrOrderBy =  " ORDER BY ssGiftCertificates.ssGCCode"

	pstrOrderBy = " Order By " & paryOrderBy(mstrOrderBy) & " " & mstrSortOrder 
	'debugprint "SummaryFilter",	pstrsqlWhere  & pstrOrderBy
	SummaryFilter = Array(pstrsqlWhere,pstrsqlHaving,pstrOrderBy)
	
End Function    'SummaryFilter

'--------------------------------------------------------------------------------------------------

Function ConvertToBoolean(vntValue)

On Error Resume Next

	vntValue = cBool(vntValue)
	If Err.number <> 0 Then vntValue = False
	ConvertToBoolean = vntValue

End Function	'ConvertToBoolean

'--------------------------------------------------------------------------------------------------
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="../SFLib/ssGiftCertificate_class.asp"-->
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/storeAdminSettings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/ssFieldValidation.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

On Error Goto 0	'added because of global error suppression in mail.asp

'**************************************************
'
'	Start Code Execution
'

mstrPageTitle = "Gift Certificate Administration"
Const cblnAutoShowTable = True

'page variables
Dim mstrAction
Dim mclsssGiftCertificate
Dim mlngID

Dim mblnShowFilter
Dim mblnShowSummary
Dim mbytSummaryTableHeight
Dim mstrShow
Dim mstrsqlWhere, mstrSortOrder,mstrOrderBy

'Display setting
Dim mbytDisplay

'Filter Elements
Dim mbytText_Filter
Dim mstrText_Filter

Dim mstrStartDate, mstrEndDate
Dim mcurLowerPrice
Dim mcurUpperPrice

Dim mbytDate_Filter
Dim mbytoptFlag_Active

'Paging Elements
Dim mlngPageCount,mlngAbsolutePage
Dim mlngMaxRecords

	mlngMaxRecords = LoadRequestValue("PageSize")
	If len(mlngMaxRecords) = 0 Then mlngMaxRecords = 50	'Set your default Maximum Records to show in summary table
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1

	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	mbytDisplay = Request.Form("optDisplay")
	If len(mbytDisplay) = 0 Then mbytDisplay = 0
	
	mstrAction = LoadRequestValue("Action")
	mlngID = LoadRequestValue("ssGCID")
	
	'Response.Write "mstrAction: " & mstrAction & "<br />" & vbcrlf
	'Response.Write "mlngID: " & mlngID & "<br />" & vbcrlf
	'Response.Flush

    Set mclsssGiftCertificate = New clsssGiftCertificate
    With mclsssGiftCertificate
		mclsssGiftCertificate.Connection = cnn
    Select Case mstrAction
        Case "New", "Update"
            .Update
            If Len(mlngID) = 0 Then mlngID = .ssGCID
            If .LoadAll(SummaryFilter) Then .LoadByID mlngID
            
            If Trim(Request.Form("SendEmail")) = "1" Then
				.validateCertificate .ssGCCode
				Call createMail("",.ssGCIssuedToEmail & "|" & "" & "|" & "" & "|" & Request.Form("emailSubject") & "|" & Request.Form("emailBody"))
				Call .updateEmailSent(True)
            End If

        Case "DeleteCertificate"
            If .LoadAll(SummaryFilter) Then 
				.LoadByID mlngID
				.Delete .ssGCCode
			End If
            .LoadAll(SummaryFilter)
			.Load .ssGCCode
        Case "View"
            If .LoadAll(SummaryFilter) Then .LoadByID mlngID
        Case "ViewByCode"
            If .LoadAll(SummaryFilter) Then .Load Request.QueryString("ssGCCode")
        Case "Activate", "Deactivate"
            .Activate mvntID, mstrAction= "Activate"
            If .LoadAll(SummaryFilter) Then .LoadByID mlngID
        Case "createRedemption"
			.Update
			.CreateRedemption True, "", .ssGCCode, 2, 0, False, "", "", ""
            If .LoadAll(SummaryFilter) Then .LoadByID mlngID
        Case "deleteRedemption"
			.deleteRedemption Request.Form("ssGCRedemption_delete")
            If .LoadAll(SummaryFilter) Then .LoadByID mlngID
        Case Else
            .LoadAll(SummaryFilter)
            If Len(.ssGCCode) > 0 Then .Load Trim(.ssGCCode)
			mblnShowFilter = False
			mblnShowSummary = True
    End Select
    
    .validateCertificate .ssGCCode
    
    'Start Output to Browswer
	Call WriteHeader("body_onload();",True)

%>
<script LANGUAGE="JavaScript">

var theDataForm;
var strDetailTitle = "<% If len(.ssGCCode) > 0 Then Response.Write "Certificate " & EncodeString(.ssGCCode,False) %>";
var blnIsDirty;

function MakeDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnReset.disabled = false;
	blnIsDirty = true;
}

function body_onload()
{
	theDataForm = document.frmData;
	blnIsDirty = false;
	document.all("spanprodName").innerHTML = strDetailTitle;

<%
If mblnShowSummary Then
	Response.Write "DisplayMainSection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplayMainSection('Filter');" & vbcrlf
Else
	Response.Write "DisplayMainSection('itemDetail');" & vbcrlf
	Response.Write "ScrollToElem('selectedSummaryItem');" & vbcrlf
End If
%>
}

function DisplayMainSection(strSection)
{

	var arySections = new Array('Filter', 'Summary', 'itemDetail');

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
	
	if (document.all("tblSummaryFunctions") != null)
	{
 		if (strSection == "Summary")
		{
			document.all("tblSummaryFunctions").style.display = "";
		}else{
			document.all("tblSummaryFunctions").style.display = "none";
		}
	}

	return(false);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewPage(theValue)
{
	theDataForm.AbsolutePage.value = theValue;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}
function View(theValue)
{
	theDataForm.ssGCID.value = theValue;
	theDataForm.Action.value = "View";
	theDataForm.submit();
	return false;
}

function validateInput(theForm)
{
	if (isEmpty(theForm.ssGCIssuedToEmail,"Please enter a value for the Issued to Email.")) {return(false);}
	
	var e;
	var partName;
	for (var i = 0; i < theForm.length; i++)
	{
		e = theForm.elements[i];
		partName = e.name.slice(0,20);
		if (partName == "ssGCRedemptionAmount")
		{
			if (!isNumeric(e,false,"Please enter an amount.")) {return(false);}
		}
	}

    return(true);
}

function btnFilter_onclick(theButton)
{

  theDataForm.Action.value = "Filter";
  theDataForm.submit();
  return(true);
}

function OpenHelp(strURL)
{
window.open(strURL,"OrderHelp","toolbar=0,location=0,directories=0,status=0,copyhistory=0,scrollbars=1,resizable=1");
}

//-->
</script>
<body onload="body_onload();">

<center>

<%= .OutputMessage %>
<form action="ssGiftCertificateAdmin.asp" id="frmData" name="frmData" onsubmit="return validateInput(this);" method="post">
  <input type="hidden" id="OrderID" name="OrderID" value="<%= .ssGCID %>">
  <input type="hidden" id="Action" name="Action" value="Update">
  <input type="hidden" id="blnShowSummary" name="blnShowSummary" value>
  <input type="hidden" id="blnShowFilter" name="blnShowFilter" value>
  <input type="hidden" id="AbsolutePage" name="AbsolutePage" value="<%= mlngAbsolutePage %>">
  <input type="hidden" id="OrderBy" name="OrderBy" value="<%= mstrOrderBy %>">
  <input type="hidden" id="SortOrder" name="SortOrder" value="<%= mstrSortOrder %>">

<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplayMainSection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your filter criteria here.">&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View items which meet the specified filter criteria">&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tditemDetail" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('itemDetail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View the selected item's detail">&nbsp;Detail&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/GiftCertificate/help_GiftCertificate.htm')" id="Button1" name="btnHelp" title="Release Version <%= mstrssAddonVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
	<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
		<colgroup align="left">
		</colgroup>
		<colgroup align="left">
		</colgroup>
		<tr>
		<td valign="top">
		<input type="radio" value="1" <% if mbytText_Filter="1" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter1"><label for="optText_Filter1">Certificate #</label><br />
		<input type="radio" value="2" <% if mbytText_Filter="2" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter2"><label for="optText_Filter2">Issued to Last Name</label><br />
		<input type="radio" value="3" <% if mbytText_Filter="3" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter3"><label for="optText_Filter3">Sending Order #</label><br />
		<input type="radio" value="4" <% if mbytText_Filter="4" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter4"><label for="optText_Filter4">Issued to email</label><br />
		<input type="radio" value="0" <% if (mbytText_Filter="0" or mbytText_Filter="") then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter0"><label for="optText_Filter0">Do 
		Not Include</label>
		<p>containing the text<br />
		<input type="text" id="Text_Filter" name="Text_Filter" size="20" value="<%= Server.HTMLEncode(mstrText_Filter) %>">
	<script language="javascript">

	function ChangeDate(theOption)
	{
		switch (theOption.value)
		{
			case "0":
				theDataForm.StartDate.value= "";
				theDataForm.EndDate.value= "";
				break;
			case "1":
				theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("d",-1,Date())) %>";
				theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
				break;
			case "2":
				theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("ww",-1,Date())) %>";
				theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
				break;
			case "3":
				theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("m",-1,Date())) %>";
				theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
				break;
			case "4":
				theDataForm.StartDate.value= "<%= FormatDateTime(DateAdd("yyyy",-1,Date())) %>";
				theDataForm.EndDate.value= "<%= FormatDateTime(Date()) %>";
				break;
			case "5":
				theDataForm.StartDate.focus();
				break;
		}
	}

	function btnReset_onclick(theButton)
	{
	var theForm = theButton.form;

		theForm.SendEmail.checked = false;
		theForm.SendEmail.disabled = false;
		theForm.btnSendEmail.disabled = false;
		
	}

	function btnNew_onclick(theButton)
	{
	var theForm = theButton.form;

		theForm.btnUpdate.value = "Save New Certificate";
		theForm.btnDelete.disabled = true;
		theForm.btnReset.disabled = false;
		
		SetDefaults(theForm);
		
	}

	function btnDelete_onclick(theButton)
	{
	var theForm = theButton.form;
	var blnConfirm;

		blnConfirm = confirm("Are you sure you wish to delete Certificate " + theForm.ssGCCode.value + "?");
		if (blnConfirm)
		{
		theForm.Action.value = "DeleteCertificate";
		theForm.submit();
		return(true);
		}
		else
		{
		return(false);
		}
	}

	function createRedemption()
	{
		theDataForm.Action.value = "createRedemption";
		theDataForm.submit();
		return false;
	}

	function deleteRedemption(strRedemptionID)
	{
		theDataForm.ssGCRedemption_delete.value = strRedemptionID;
		theDataForm.Action.value = "deleteRedemption";
		theDataForm.submit();
		return false;
	}

	function SetDefaults(theForm)
	{

		theForm.ssGCID.value = "";
		theForm.ssGCCode.value = "---";
		theForm.ssGCExpiresOn.value = "";
		theForm.ssGCSingleUse.checked = false;
		theForm.ssGCCustomerID.value = "";

		theForm.SendEmail.checked = false;
		theForm.SendEmail.disabled = true;
		theForm.btnSendEmail.disabled = true;
		theForm.ssGCElectronic.checked = false;

		theForm.ssGCToName.value = "";
		theForm.ssGCIssuedToEmail.value = "";
		theForm.ssGCFromName.value = "";
		theForm.ssGCFromEmail.value = "";
		theForm.ssGCMessage.value = "";
	    
		var MTPTable = document.all("tblRedemptions");
		//row 0 is header
		//row 1-3 is data to keep
		//row n is footer
		//need to delete all rows between n-1 and 3
		for (var i=MTPTable.rows.length-2; i>3; i--)
		{
		MTPTable.deleteRow(i);
		}

		var lngID = theForm.ssGCRedemptionID.value;
		var theElement;
		var strElementName;
		//var strElementPrefix;
		//if (lngID != ""){strElementPrefix = strElementName + "."}else{strElementPrefix = strElementName}
		//strElementPrefix = strElementName + "."
		
		strElementName = "ssGCRedemptionAmount."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.value = "";
		
		strElementName = "ssGCRedemptionType."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		SetSelect(theElement,2)
	    
		strElementName = "ssGCRedemptionOrderID."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.value = "";
	    
		strElementName = "ssGCRedemptionActive."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.checked = false;

		strElementName = "ssGCRedemptionCreatedOn."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.value = "";

		strElementName = "ssGCRedemptionInternalNotes."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.value = "";

		strElementName = "ssGCRedemptionExternalNotes."
		theElement = document.all(strElementName + lngID);
		theElement.name = strElementName;
		theElement.id = strElementName;
		theElement.value = "";

		theForm.ssGCRedemptionID.value = "";
	    
		theForm.ssGCCode.select();
		theForm.ssGCCode.focus();
	    
	return(true);
	}

	</script>

		<td valign="top">Show Certificates Issued:<br />
		<input type="radio" value="1" <% if mbytDate_Filter="1" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter1"><label for="optDate_Filter1">Day</label>&nbsp;
		<input type="radio" value="2" <% if mbytDate_Filter="2" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter2"><label for="optDate_Filter2">Week</label>&nbsp;
		<input type="radio" value="3" <% if mbytDate_Filter="3" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter3"><label for="optDate_Filter3">Month</label>&nbsp;
		<input type="radio" value="4" <% if mbytDate_Filter="4" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter4"><label for="optDate_Filter4">Year</label>&nbsp;
		<input type="radio" value="5" <% if mbytDate_Filter="5" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter5"><label for="optDate_Filter5">Custom</label>&nbsp;
		<input type="radio" value="0" <% if mbytDate_Filter="0" then Response.Write "Checked" %> name="optDate_Filter" onclick="ChangeDate(this);" ID="optDate_Filter0"><label for="optDate_Filter0">All</label>&nbsp;<br />
		<label for="StartDate">Start Date:&nbsp;</label><input id="StartDate" name="StartDate" Value="<%= mstrStartDate %>" size="20">
		<a HREF="javascript:doNothing()" title="Select start date" onClick="setDateField(document.frmData.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER="0"></a><br />
		<label for="EndDate">&nbsp;&nbsp;End Date:&nbsp;</label><input id="EndDate" name="EndDate" Value="<%= mstrEndDate %>" size="20">
		<a HREF="javascript:doNothing()" title="Select end date" onClick="setDateField(document.frmData.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER="0"></a>

			<p>Totals between<br />
			<input type="text" id="LowerPrice" name="LowerPrice" size="5" value="<%= mcurLowerPrice %>" maxlength=15>
			And
			<input type="text" id="UpperPrice" name="UpperPrice" size="5" value="<%= mcurUpperPrice %>" maxlength=15></p>
		</td>
		<td valign="top">
		<table cellpadding="0" cellspacing="0" border="0" ID="Table2">
			<caption align="left"><font size="-1">Show Certificates that are:</font></caption>
			<tr>
			<td>
				<input type="radio" value="1" <% if mbytoptFlag_Active="1" then Response.Write "Checked" %> name="optFlag_Active" ID="optFlag_Active1"><label for="optFlag_Active1"><font size="-1">Active</font></label><td>
				&nbsp;<input type="radio" value="2" <% if mbytoptFlag_Active="2" then Response.Write "Checked" %> name="optFlag_Active" ID="optFlag_Active2"><label for="optFlag_Active2"><font size="-1">Inactive</font></label><td>
				&nbsp;<input type="radio" value="0" <% if (mbytoptFlag_Active="0" or mbytoptFlag_Active="") then Response.Write "Checked" %> name="optFlag_Active" ID="optFlag_Active0"><label for="optFlag_Active0"><font size="-1">Both</font></label></tr>
		</table>
		<td valign="top">
		<input class="butn" id="btnFilter" name="btnFilter" type="button" value="Apply Filter" onclick="btnFilter_onclick(this);"><p>&nbsp;
		</td> 
		</tr>
	</table>
	<% Call OutputSummary(.SummaryRecordset) %>
	<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblitemDetail">
		<colgroup align="left" width="25%">
		</colgroup>
		<colgroup align="left" width="75%">
		</colgroup>
		<tr class="tblhdr">
		<th align="center"><span id="spanprodName"></span>&nbsp;</th>
		</tr>
		<tr>
		<td>
		<% Call ShowGiftCertificateDetail(mclsssGiftCertificate) %>
		<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" rules="none" id="tblFooter">
			<tr>
			<td>&nbsp;<td>
			<a href="" onclick="window.open('../../../viewCertificate.asp?Certificate=<%= mclsssGiftCertificate.ssGCCode %>','viewCertificate','toolbar=0,location=0,directories=0,status=0,copyhistory=0,scrollbars=0, width=800, height=600, screenX=200, screenY=300,');return false;"><font size="-1">View Certificate</font></a>&nbsp;
			<input class="butn" type="button" name="btnHelp" ID="btnHelp" value="?" onclick="OpenHelp('ssHelpFiles/GiftCertificate/help_GiftCertificate.htm');">&nbsp;
			<input class="butn" type="reset" name="btnReset" ID="btnReset" value="Reset" onclick="btnReset_onclick(this);">&nbsp;&nbsp;
			<input class="butn" type="button" name="btnNew" ID="btnNew" value="New" onclick="return btnNew_onclick(this)">&nbsp;
			<input class="butn" type="button" name="btnDelete" ID="btnDelete" value="Delete" onclick="return btnDelete_onclick(this)">
			<input class="butn" type="submit" name="btnUpdate" ID="btnUpdate" value="Save Changes">
			</tr>
		</table>
		</tr>
	</table>
  <span id="spantempFile" style="display:none">&nbsp;</span></p>
	</td>
  </tr>
</table>

</form>
</center>

</body>

</html>
<%
    End With

    Set cnn = Nothing
    Set mclsssGiftCertificate = Nothing
    Response.Flush

'******************************************************************************************************************************************************************

Sub OutputSummary(objRS)

	'On Error Resume Next

	Dim i
	Dim aSortHeader(4,3)
	Dim pstrOrderBy, pstrSortOrder, pstrTempSort
	Dim pstrTitle
	Dim pstrSelect, pstrHighlight
	Dim pstrID
	Dim pblnSelected
	Dim pblnClosed

		With Response

			If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
				pstrTempSort = "descending"
				pstrSortOrder = "ASC"
			Else
				pstrTempSort = "ascending"
				pstrSortOrder = "DESC"
			End If
			
			aSortHeader(1,0) = "Sort by Certificate in " & pstrTempSort & " order"
			aSortHeader(2,0) = "Sort by Issued To in " & pstrTempSort & " order"
			aSortHeader(3,0) = "Sort by Amount in " & pstrTempSort & " order"
			aSortHeader(4,0) = "Sort by Issue Date in " & pstrTempSort & " order"
				
			aSortHeader(1,1) = "Certificate"
			aSortHeader(2,1) = "Issued To"
			aSortHeader(3,1) = "Amount"
			aSortHeader(4,1) = "Issue Date"

			'column header widths
			aSortHeader(0,2) = 2
			aSortHeader(1,2) = 25
			aSortHeader(2,2) = 25
			aSortHeader(3,2) = 20
			aSortHeader(4,2) = 25

			'scrolling column widths
			aSortHeader(0,3) = 2
			aSortHeader(1,3) = 26
			aSortHeader(2,3) = 26
			aSortHeader(3,3) = 22
			aSortHeader(4,3) = 24

			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblSummary' rules='none'>" & vbcrlf	'
			For i = 0 to UBound(aSortHeader)
				.Write "    <colgroup align='left' width='" & aSortHeader(i,2) & "%'>" & vbcrlf
			Next 'i
			
			if len(mstrOrderBy) > 0 Then
				pstrOrderBy = mstrOrderBy
			Else
				pstrOrderBy = "1"
			End If
		
			.Write "    <tr class='tblhdr'>" & vbcrlf
			.Write "      <th>&nbsp;</th>" & vbcrlf
			For i = 1 to UBound(aSortHeader)
				If cInt(pstrOrderBy) = i Then
					If (pstrSortOrder = "ASC") Then
						.Write "      <th style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></th>" & vbCrLf
					Else
						.Write "      <th style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></th>" & vbCrLf
					End If
				Else
				    .Write "      <th style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
									" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
									" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</th>" & vbCrLf
				End If
			Next 'i
			.Write "      </tr>" & vbcrlf
	' 
			'Now for the summary table contents
			.Write "      <tr><td colspan='" & UBound(aSortHeader) + 1 & "'>" & vbcrlf
			.Write "        <div name='divSummary' style='height:400; overflow:scroll;'>" & vbcrlf	'
			.Write "        <table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none'" _
					 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
					 & ">" & vbcrlf
					 
			For i = 0 to UBound(aSortHeader)
				.Write "          <colgroup align='left' width='" & aSortHeader(i,3) & "%'>" & vbcrlf
			Next 'i

			If objRS.State = 1 Then
			If objRS.RecordCount > 0 Then
				objRS.MoveFirst

				'Need to calculate current recordset page and upper bound to loop through
				dim plnguBound, plnglbound, pstrDisplay

				If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
				If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
				plnglbound = (mlngAbsolutePage - 1) * objRS.PageSize + 1
				plnguBound = mlngAbsolutePage * objRS.PageSize

				If plnguBound > objRS.RecordCount Then plnguBound = objRS.RecordCount
				objRS.AbsolutePosition = plnglbound
				
				Dim pcurValue
				Dim pcurTempValue
				Dim pstrOutput
				
				pcurValue = 0
				For i = plnglbound To plnguBound
					pcurTempValue = objRS.Fields("CertificateValue").Value
					'If Not isNumeric(pcurTempValue) Then pcurTempValue = 0
					If Not isNull(objRS.Fields("CertificateValue").Value) Then
						pcurTempValue = CDbl(Trim(objRS.Fields("CertificateValue").Value))
					Else
						pcurTempValue = 0
					End If
					pcurValue = pcurValue + pcurTempValue
					If pstrID <> trim(objRS("ssGCCode")) Or True Then
						If Len(pstrOutput) <> 0 Then
							'pstrOutput = Replace(pstrOutput,"<<certificateValue>>",FormatCurrency(pcurValue,2))
	        				Response.Write pstrOutput
							pstrOutput = ""
						End If
						pstrID = trim(objRS("ssGCCode"))
						pstrTitle = "Click to view " & objRS("ssGCCode")
						pstrHighlight = "title='" & pstrTitle & "' " _
									& "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
									& "onmouseout='doMouseOutRow(this); ClearTitle();' " _
									& "onmousedown='View(" & chr(34) & objRS.Fields("ssGCID").Value & chr(34) & ");'"

						pstrSelect = "title='" & pstrTitle & "' " _
								& "onmouseover='DisplayTitle(this);' " _
								& "onmouseout='ClearTitle();' " _
								& "onmousedown='View(" & chr(34) & objRS.Fields("ssGCID").Value & chr(34) & ");'"
					
						pblnSelected = (pstrID = mclsssGiftCertificate.ssGCCode)

						If pblnSelected Then
							pstrOutput = pstrOutput & "<tr class='Selected' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);'>" & vbcrlf
							pstrOutput = pstrOutput & " <td>&nbsp;</td>"
							'.Write " <td><input type=checkbox value=" & Chr(34) & Server.HTMLEncode(pstrID) & Chr(34) & " checked></td>" & vbcrlf
						Else
							If pblnClosed Then
								pstrOutput = pstrOutput & " <tr class='Inactive' " & pstrHighlight & ">" & vbcrlf
							Else
								pstrOutput = pstrOutput & " <tr class='Active' " & pstrHighlight & ">" & vbcrlf
							End If
							pstrOutput = pstrOutput & " <td " & pstrSelect & ">&nbsp;</td>" & vbcrlf
						End If

	        			pstrOutput = pstrOutput & "<td align=left>&nbsp;&nbsp;" & pstrID & "&nbsp;</td>" & vbcrlf
	        			pstrOutput = pstrOutput & "<td align=left>" & objRS.Fields("ssGCIssuedToEmail").Value & "&nbsp;</td>" & vbcrlf
	        			pstrOutput = pstrOutput & "<td align=left>" & FormatCurrency(pcurTempValue,2) & "&nbsp;</td>" & vbcrlf

						If isNull(objRS.Fields("ssGCCreatedOn").Value) Then
		        			pstrOutput = pstrOutput & "<td align=left>&nbsp;</td>" & vbcrlf
		        		Else
		        			pstrOutput = pstrOutput & "<td align=left>" & FormatDateTime(objRS.Fields("ssGCCreatedOn").Value,2) & "&nbsp;</td>" & vbcrlf
		        		End If
		        	
	        			pstrOutput = pstrOutput & "</tr>" & vbcrlf
	        		End If
					objRS.MoveNext
				Next
				If Len(pstrOutput) <> 0 Then
					pstrOutput = Replace(pstrOutput,"<<certificateValue>>",FormatCurrency(pcurValue,2))
	        		Response.Write pstrOutput
					pstrOutput = ""
				End If
			Else
					.Write "          <tr><td align=center COLSPAN=" & UBound(aSortHeader) + 1 & "><h3>There are no Gift Certificates</h3></td></tr>" & vbcrlf
			End If	'objRS.RecordCount > 0
			End If	'objRS.State = 1
			.Write "        </TABLE></div>" & vbcrlf
			.Write "      </td></tr>" & vbcrlf

			'Write the paging routine
			.Write "        <tr class='tblhdr'><th COLSPAN='" & UBound(aSortHeader) + 1 & "' align=center>"
			
			If objRS.State = 1 Then
			If objRS.RecordCount = 0 Then
				.Write "No Gift Certificates match your search criteria"
			Elseif objRS.RecordCount = 1 Then
				.Write "1 Gift Certificate matches your search criteria"
			Else 
				.Write objRS.RecordCount & " Gift Certificates match your search criteria with a value of " & FormatCurrency(pcurValue,2) & "<br />"

				dim pstrCheck
				pstrCheck = "return isInteger(this, true, ""Please enter a positive integer for the recordset page size."");"
				.Write "Show&nbsp;<input type='Text' id='PageSize' name='PageSize' value='" & objRS.PageSize & "' maxlength='4' size='4' style='text-align:center;' onblur='" & pstrCheck & "'>&nbsp;records at a time.&nbsp;&nbsp;"
				For i=1 to mlngPageCount
					plnglbound = (i-1) * mlngMaxRecords + 1
					plnguBound = i * mlngMaxRecords
					if plnguBound > objRS.RecordCount Then plnguBound = objRS.RecordCount
					pstrDisplay = plnglbound & " - " & plnguBound & "&nbsp;"
					If i = cInt(mlngAbsolutePage) Then
						Response.Write pstrDisplay
					Else
						Response.Write "<a href='#' onclick='return ViewPage(" & i & ");'>" & pstrDisplay & "</a>&nbsp;"
					End If
				Next
			End If
			Else
				.Write "<font size=+2 color=red>Error opening database</font>"
			End If
			.Write "</th></TR>" & vbcrlf
			.Write "      </TABLE>" & vbcrlf
		End With
		
End Sub      'OutputSummary

'******************************************************************************************************************************************************************

Sub ShowGiftCertificateDetail(objclsssGiftCertificate)

Dim pstrCustomerLink

With objclsssGiftCertificate

If isObject(.Recordset) Then 
If .Recordset.RecordCount > 0 Then 
	.Recordset.MoveFirst 
	If Len(.Recordset.Fields("ssGCCustomerID").Value & "") > 0 Then 
		If cblnSQLDatabase Then
			pstrCustomerLink = .Recordset.Fields("custLastName").Value & ", " & .Recordset.Fields("custFirstName").Value
		Else
			pstrCustomerLink = .Recordset.Fields("sfCustomers.custLastName").Value & ", " & .Recordset.Fields("sfCustomers.custFirstName").Value
		End If
		
		'Now make it a link to Customer Manager
		pstrCustomerLink = "<a href=""sfCustomerAdmin.asp?Action=viewItem&ViewID=" & .Recordset.Fields("ssGCCustomerID").Value & """ title=""View customer record"" target=_blank>" & pstrCustomerLink & "</a>"
	End If
End If
End If
%>
<input type="hidden" name="ssGCID" id="ssGCID" value="<%= .ssGCID %>">
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblDetail">
  <tr>
    <td width="100%">
     <table border="0" cellspacing="0" cellpadding="3" width="100%" ID="tblOrderDetailSummary">
       <tr>
         <td class="label">&nbsp;<label for="ssGCCode">Code:</label></td>
         <td>&nbsp;<input id="ssGCCode" name="ssGCCode" Value="<%= .ssGCCode %>" maxlength="20" size="20"></td>
       </tr>
       <tr>
         <td class="label">&nbsp;<label for="ssGCCustomerID">Customer:</label></td>
         <td>&nbsp;<input id="ssGCCustomerID" name="ssGCCustomerID" Value="<%= .ssGCCustomerID %>" size="6" title="This is the customer id as found in the sfCustomers table">&nbsp;<%= pstrCustomerLink %></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>
           <table border="1" cellpadding="2" cellspacing="0">
				<tr>
					<td class="label">&nbsp;<label id="lblssGCFromName" for="ssGCFromName">From Name:</label></td>
					<td>&nbsp;<input id="ssGCFromName" name="ssGCFromName" Value="<%= .ssGCFromName %>" maxlength="255" size="60"></td>
				</tr>
				<tr>
					<td class="label">&nbsp;<label id="lblssGCFromEmail" for="ssGCFromEmail">From Email:</label></td>
					<td>&nbsp;<input id="ssGCFromEmail" name="ssGCFromEmail" Value="<%= .ssGCFromEmail %>" maxlength="255" size="60"></td>
				</tr>
				<tr>
					<td class="label">&nbsp;<label id="lblssGCToName" for="ssGCToName">To Name:</label></td>
					<td>&nbsp;<input id="ssGCToName" name="ssGCToName" Value="<%= .ssGCToName %>" maxlength="255" size="60"></td>
				</tr>
				<tr>
					<td class="label">&nbsp;<label id="lbl" for="ssGCIssuedToEmail">To Email:</label></td>
					<td>&nbsp;<input id="ssGCIssuedToEmail" name="ssGCIssuedToEmail" Value="<%= .ssGCIssuedToEmail %>" maxlength="255" size="60"></td>
				</tr>
				<tr>
					<td class="label">&nbsp;<label id="lblssGCMessage" for="ssGCMessage">Message:</label></td>
					<td>&nbsp;<textarea rows="9" name="ssGCMessage" ID="ssGCMessage" cols="70"><%= .ssGCMessage %></textarea></td>
				</tr>
           </table>
         </td>
       </tr>

<%
Dim mstremailFile
Dim pstrEmailBody
Dim pstrEmailSubject

Call LoadEmails(mstremailFile, pstrEmailSubject, pstrEmailBody, objclsssGiftCertificate)

%>
 <tr><td colspan=2 align=left>
	<div id="divEmail" style="position:absolute; display:none">
<table border="3" cellspacing="0" cellpadding="0" bgcolor="white" id="tblEmail" style="border-style:outset; border-color:steelblue;"><tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="3" ID="Table8">
  <tr>
    <td align="right">Select an email template:</td>
    <td>
      <script language="javascript">
      function changeEmailTemplate(theSelect)
      {
      theSelect.form.emailSubject.value  = document.all("enEmail_Subject" + theSelect.selectedIndex).value;
      
      theSelect.form.emailBody.value  = document.all("enEmail_Body" + theSelect.selectedIndex).value;
      }
      </script>
      <% For i = 0 To UBound(maryEmails) %>
      <input type="hidden" name="enEmail_Subject<%= i %>" id="enEmail_Subject<%= i %>" value="<%= maryEmails(i)(enEmail_Subject) %>">
      <input type="hidden" name="enEmail_Body<%= i %>" id="enEmail_Body<%= i %>" value="<%= maryEmails(i)(enEmail_Body) %>">
      <% Next 'i %>
      <select name="emailFile" ID="emailFile" onchange="changeEmailTemplate(this); return false;">
      <% For i = 0 To UBound(maryEmails) %>
      <%   If CBool(mstremailFile = maryEmails(i)(enEmail_FileName)) Or CBool(Len(mstremailFile)=0 And (i = UBound(maryEmails))) Then  %>
      <%	pstrEmailSubject = maryEmails(i)(enEmail_Subject) %>
      <%	pstrEmailBody = maryEmails(i)(enEmail_Body) %>
		<option selected><%= maryEmails(i)(enEmail_FileName) %></option>
      <%   Else %>
		<option><%= maryEmails(i)(enEmail_FileName) %></option>
      <%   End If %>
      <% Next 'i %>
      </select>
    </td>
  </tr>
  <tr>
    <td align="right">
      <p>To:</td>
    <td><input type="text" name="emailTo" ID="emailTo" size="75" VALUE="<%= .ssGCIssuedToEmail %>"></td>
  </tr>
  <tr>
    <td align="right">Subject:</td>
    <td><input type="text" name="emailSubject" ID="emailSubject" size="75" VALUE="<%= pstrEmailSubject %>"></td>
  </tr>
  <tr>
    <td align="right">Body:</td>
    <td><textarea rows="9" name="emailBody" ID="emailBody" cols="70"><%= pstrEmailBody %></textarea></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input class='butn' type="button" value="Send" name="B1" ID="B1" onclick="document.frmData.SendEmail.checked=true; document.all('divEmail').style.display = 'none';">&nbsp;
        <input class='butn' type="button" value="Cancel" name="Cancel" ID="Cancel" onclick='document.all("divEmail").style.display = "none";'></td>
		<input type=hidden id=StockEmail name=StockEmail value=1>
  </tr>
</table>
</td></tr></table>
</div>
</td></tr>
        <tr>
          <td>&nbsp;&nbsp;</td>
          <td>
		    <input type="checkbox" id="SendEmail" name="SendEmail" value="1" title="Check this box to send the email. Email will be sent when Save Changes is pressed"><label for"SendEmail">Send Email</label>&nbsp;
		    <input type="button" class="butn" id="btnSendEmail" name="btnSendEmail" value="Edit Email Text" onclick='document.all("divEmail").style.display = ""; document.frmData.StockEmail.value=0; document.frmData.emailBody.focus();'>
          </td>
        </tr>
       <tr>
         <td class="label">&nbsp;</td>
         <td><input id="ssGCElectronic" name="ssGCElectronic" type="checkbox" <% If .ssGCElectronic Then Response.Write "Checked" %> value="ON">&nbsp;<label for="ssGCElectronic">Email Sent</label></td>
       </tr>
       <tr>
         <td class="label">&nbsp;</td>
         <td><input id="ssGCSingleUse" name="ssGCSingleUse" type="checkbox" <% If .ssGCSingleUse Then Response.Write "Checked" %> value="ON" title="Check this box to make certificate single use.">&nbsp;<label for="ssGCSingleUse">Single use</label></td>
       </tr>
       <tr>
         <td class="label">&nbsp;<label id="lblssGCExpiresOn" for="ssGCExpiresOn" title="Double click to set expiration date to one year from today." ondblclick="theDataForm.ssGCExpiresOn.value='<%= DateAdd("y",1,Date()) %>';">Expires On: </label>
         <td>&nbsp;<input id="ssGCExpiresOn" name="ssGCExpiresOn" Value="<%= .ssGCExpiresOn %>" size="20">
         <a HREF="javascript:doNothing()" title="Select start date" onClick="setDateField(document.frmData.ssGCExpiresOn); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
         <img SRC="images/calendar.gif" BORDER="0"></a></td>
       </tr>
       <tr>
         <td class="label">&nbsp;<label id="lblssGCCreatedOn" for="ssGCCreatedOn">Created On:</label></td>
         <td>&nbsp;<%= .ssGCCreatedOn %></td>
       </tr>
       <tr>
         <td class="label">&nbsp;<label id="lblssGCModifiedOn" for="ssGCModifiedOn">Modified On:</label></td>
         <td>&nbsp;<%= .ssGCModifiedOn %></td>
       </tr>
       <tr>
         <td class="label">&nbsp;<td>&nbsp;</tr>
       <tr>
         <td></td>
         <td>
          <input type="hidden" name="ssGCRedemption_delete" id="ssGCRedemption_delete" value="">
          <table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="tblRedemptions">
            <tr>
              <th>&nbsp;</th>
              <th>Amount</th>
              <th>Type</th>
              <th>Order #</th>
              <th>Active</th>
              <th>Created On</th>
            </tr>
            <% 
			If isObject(.Recordset) Then 
            If .Recordset.RecordCount > 0 Then 
				.Recordset.MoveFirst 
				.Recordset.Filter = "ssGCID=" & .ssGCID
            Dim i
            Dim plngRedemptionID
            For i = 1 To .Recordset.Recordcount
				plngRedemptionID = .Recordset.Fields("ssGCRedemptionID").Value
            %>
            
            <tr>
              <td align="center">
                <input type="hidden" name="ssGCRedemptionID" id="ssGCRedemptionID" value="<%= plngRedemptionID %>"><%= i %>
                <a href="" onclick="deleteRedemption('<%= plngRedemptionID %>'); return false;" title="Delete this">x</a>
              </td>
              <td align="center"><input  id="ssGCRedemptionAmount.<%= plngRedemptionID %>"    name="ssGCRedemptionAmount.<%= plngRedemptionID %>" value="<%= .Recordset.Fields("ssGCRedemptionAmount").Value %>" size="20"></td>
              <td align="center"><select id="ssGCRedemptionType.<%= plngRedemptionID %>"      name="ssGCRedemptionType.<%= plngRedemptionID %>"><% Call WriteRedemptionTypeCombo(.Recordset.Fields("ssGCRedemptionType").Value) %></select></td>
              <td align="center"><input  id="ssGCRedemptionOrderID.<%= plngRedemptionID %>"   name="ssGCRedemptionOrderID.<%= plngRedemptionID %>" value="<%= .Recordset.Fields("ssGCRedemptionOrderID").Value %>" size="20"></td>
              <td align="center"><input  id="ssGCRedemptionActive.<%= plngRedemptionID %>"    name="ssGCRedemptionActive.<%= plngRedemptionID %>" type="checkbox" <% If .Recordset.Fields("ssGCRedemptionActive").Value Then Response.Write "Checked" %> value="ON"></td>
              <td align="center"><input  id="ssGCRedemptionCreatedOn.<%= plngRedemptionID %>" name="ssGCRedemptionCreatedOn.<%= plngRedemptionID %>" value="<%= .Recordset.Fields("ssGCRedemptionCreatedOn").Value %>" size="20" disabled></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="2"><font size="-1">External Notes</font><br /><textarea id="ssGCRedemptionExternalNotes.<%= plngRedemptionID %>" name="ssGCRedemptionExternalNotes.<%= plngRedemptionID %>" rows="2" cols="40"><%= .Recordset.Fields("ssGCRedemptionExternalNotes").Value %></textarea></td>
              <td colspan="3"><font size="-1">Internal Notes</font><br /><textarea id="ssGCRedemptionInternalNotes.<%= plngRedemptionID %>" name="ssGCRedemptionInternalNotes.<%= plngRedemptionID %>" rows="2" cols="40"><%= .Recordset.Fields("ssGCRedemptionInternalNotes").Value %></textarea></td>
            </tr>
            <tr><td colspan="6"><hr></td></tr>
            <% 
             .Recordset.MoveNext 
             Next 'i
            .Recordset.Filter = ""
            %>
            <tr><td colspan="6" align="center"><a href="" onclick="createRedemption(); return false;" title="Create new entry">New</a></td></tr>
            <%
            Else
            %>
            <tr>
              <td align="center"><input type="hidden" name="ssGCRedemptionID" id="Hidden1" value="<%= plngRedemptionID %>"></td>
              <td align="center"><input  id="ssGCRedemptionAmount."    name="ssGCRedemptionAmount." value="" size="20"></td>
              <td align="center"><select id="ssGCRedemptionType."      name="ssGCRedemptionType."><% Call WriteRedemptionTypeCombo(2) %></select></td>
              <td align="center"><input  id="ssGCRedemptionOrderID."   name="ssGCRedemptionOrderID." value="" size="20"></td>
              <td align="center"><input  id="ssGCRedemptionActive."    name="ssGCRedemptionActive." type="checkbox" value="ON"></td>
              <td align="center"><input  id="ssGCRedemptionCreatedOn." name="ssGCRedemptionCreatedOn." value="" size="20" disabled></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="2"><font size="-1">External Notes</font><br /><textarea id="ssGCRedemptionExternalNotes." name="ssGCRedemptionExternalNotes." rows="2" cols="40"></textarea></td>
              <td colspan="3"><font size="-1">Internal Notes</font><br /><textarea id="ssGCRedemptionInternalNotes." name="ssGCRedemptionInternalNotes." rows="2" cols="40"></textarea></td>
            </tr>
            <tr><td colspan="6"><hr></td></tr>
            <%
			End If
			End If
            %>
          </table>
          </td>
       </tr>
     </table>
  </tr>
</table>
<% 
End With	'objclsssGiftCertificate

End Sub 'ShowGiftCertificateDetail

'***********************************************************************************************

Sub WriteRedemptionTypeCombo(lngID)

Dim plngTempID
Dim i

	If Len(lngID & "") = 0 Then
		plngTempID = 2
	Else
		plngTempID = CLng(lngID)
	End If

	For i = 1 To UBound(maryGCRedemptionTypes)
		If (plngTempID = i) Then
			Response.Write "<option value=" & i & " selected>" & maryGCRedemptionTypes(i) & "</option>"
		Else
			Response.Write "<option value=" & i & ">" & maryGCRedemptionTypes(i) & "</option>"
		End If
	Next 'i

End Sub	'WriteRedemptionTypeCombo

'***********************************************************************************************

	Sub GetEmailBody(strEmailSubject, strEmailBody, objclsssGiftCertificate, blnComplete)

	Dim p_strSubject
	Dim p_strBody
	Dim pstrTrackingLink
	Dim pstrCustName, pstrShipAddr

	'On Error Resume Next

		Call LoadEmailFile(p_strSubject, p_strBody)
		
		'now replace the constants
		strEmailSubject = customReplacements(strEmailSubject, objclsssGiftCertificate)
		strEmailBody = customReplacements(strEmailBody, objclsssGiftCertificate)

	End Sub	'GetEmailBody

'***********************************************************************************************

	Sub LoadEmailFile(byRef strEmailSubject, byRef strEmailBody)

	Dim pobjFSO
	Dim MyFile
	Dim pstrTempLine
	Dim pstrFilePath
	Dim p_strSubject
	Dim p_strBody

	'On Error Resume Next

		pstrFilePath = Request.ServerVariables("PATH_TRANSLATED")
		pstrFilePath = Replace(Lcase(pstrFilePath), "ssgiftcertificateadmin.asp", "ssSamples/ssGiftCertificateEmail.txt")

		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		Set MyFile =pobjFSO.OpenTextFile(pstrFilePath,1,True)

		p_strSubject = MyFile.ReadLine
		pstrTempLine = MyFile.ReadLine	'garbage line
		pstrTempLine = MyFile.ReadLine & vbcrlf
		Do While pstrTempLine <> "// DO NOT REMOVE THIS LINE //" AND NOT MyFile.AtEndOfStream
			p_strBody = p_strBody & pstrTempLine & vbcrlf
			pstrTempLine = MyFile.ReadLine
		Loop
		
		strEmailSubject = p_strSubject
		strEmailBody = p_strBody

		MyFile.Close
		Set MyFile = Nothing
		Set pobjFSO = Nothing

	End Sub	'LoadEmailFile

	'***********************************************************************************************

	Sub closeObj(objItem)
		ReleaseObject objItem
	End Sub

	'***********************************************************************************************

	Sub LoadEmails(ByVal strFileName, ByRef strEmailSubject, ByRef strEmailBody, ByRef objclsssGiftCertificate)

	Dim i

	'On Error Resume Next

		Call LoadEmailFiles(maryEmails)
		For i = 0 To UBound(maryEmails)
			maryEmails(i)(enEmail_Subject) = customReplacements(maryEmails(i)(enEmail_Subject), objclsssGiftCertificate)
			maryEmails(i)(enEmail_Body) = customReplacements(maryEmails(i)(enEmail_Body), objclsssGiftCertificate)

			If CBool(maryEmails(i)(enEmail_FileName) = strFileName) Or CBool((i = UBound(maryEmails)) And (Len(strFileName) = 0)) Then
				strEmailSubject = customReplacements(maryEmails(i)(enEmail_Subject), objclsssGiftCertificate)
				strEmailBody = customReplacements(maryEmails(i)(enEmail_Body), objclsssGiftCertificate)
			End If
		Next 'i

	End Sub	'LoadEmails

	'***********************************************************************************************

	Dim maryEmails

	'***********************************************************************************************
	
	Sub LoadEmailFiles(ByRef aryEmails)

	Dim pobjFSO
	Dim pobjFolder, pobjFiles
	Dim i
	Dim MyFile
	Dim pstrTempLine
	Dim pstrFilePath
	Dim p_strSubject
	Dim p_strBody
	
	On Error Resume Next

		pstrFilePath = Request.ServerVariables("PATH_TRANSLATED")
		pstrFilePath = Replace(Lcase(pstrFilePath),"ssgiftcertificateadmin.asp","")
		pstrFilePath = pstrFilePath & "emailTemplates/"

		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		i = 0
		Set pobjFolder = pobjFSO.GetFolder(pstrFilePath)
		Set pobjFiles = pobjFolder.Files
		ReDim aryEmails(pobjFiles.Count - 1)
		For Each MyFile In pobjFiles
			p_strBody = ""
			aryEmails(i) = Array("fileName", "subject", "body")
			
			aryEmails(i)(enEmail_FileName) = MyFile.Name
			Set MyFile =pobjFSO.OpenTextFile(pstrFilePath & MyFile.Name,1,True)

			p_strSubject = MyFile.ReadLine
			pstrTempLine = MyFile.ReadLine	'garbage line
			pstrTempLine = MyFile.ReadLine & vbcrlf
			Do While pstrTempLine <> "// DO NOT REMOVE THIS LINE //" AND NOT MyFile.AtEndOfStream
				p_strBody = p_strBody & pstrTempLine & vbcrlf
				pstrTempLine = MyFile.ReadLine
			Loop
			
			aryEmails(i)(enEmail_Subject) = p_strSubject
			aryEmails(i)(enEmail_Body) = p_strBody

			MyFile.Close
			Set MyFile = Nothing
			
			i = i + 1
		Next 'MyFile
		Set pobjFiles = Nothing
		Set pobjFolder = Nothing
		
		Set pobjFSO = Nothing

	End Sub	'LoadEmailFiles

	'***********************************************************************************************
%>