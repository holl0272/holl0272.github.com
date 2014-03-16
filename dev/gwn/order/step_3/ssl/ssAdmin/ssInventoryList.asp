<%
Option Explicit
Response.Buffer = False
Server.ScriptTimeout = 900

'********************************************************************************
'*   Customer Manager for StoreFront 5.0                                        *
'*   Release Version:	2.00.004		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		March 16, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.004 (March 16, 2005)													*
'*   - Enhancement - added tabbed interface										*
'*                                                                              *
'*   2.00.003 (January 14, 2004)                                                *
'*   - Bug fix - update routine modified to use nulls instead of empty values   *
'*                                                                              *
'*   2.00.002 (November 6, 2003)                                                *
'*   - Added Pricing Level Manager support                                      *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'NONE
	
'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************


Class clsItem
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pblnError
Private pstrMessage
Private pobjRS

'***********************************************************************************************

Private Sub class_Initialize()

End Sub

Private Sub class_Terminate()

Dim i

    On Error Resume Next
	Call ReleaseObject(pobjRS)

End Sub

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

'***********************************************************************************************

Public Property Get Records()
    If isObject(pobjRS) Then Set Records = pobjRS
End Property

'***********************************************************************************************

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

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	pobjRS = server.CreateObject("adodb.recordset")
	With pobjRS
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
        
		pstrSQL = "SELECT sfInventory.invenId, sfInventory.invenProdId, sfProducts.prodName, sfInventory.invenAttName, sfInventory.invenInStock" _
				& " FROM sfProducts RIGHT OUTER JOIN sfInventory ON sfProducts.prodID = sfInventory.invenProdId" _
				& " ORDER BY sfInventory.invenId"

		'Response.Write "pstrSQL: " & pstrSQL & "<br />"        
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If

		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
		Dim plnglbound
		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
		plnglbound = (mlngAbsolutePage - 1) * pobjRS.PageSize + 1
		If Not pobjRS.EOF Then pobjRS.AbsolutePosition = plnglbound

	End With

    Load = (Not pobjRS.EOF)

End Function    'Load

'******************************************************************************************************************************************************************

End Class   'clsItem

'**********************************************************

Function getCookie_SessionID()
	getCookie_SessionID = Request.Cookies("sfOrder")("SessionID")
End Function
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'**********************************************************
'*	Functions
'**********************************************************

'Sub WriteFormOpener
'Sub WriteItemDetail
'Sub WriteCustomTable
'Function CustomDisplayText(byVal lngPos, byRef ary)
'Sub CustomOutput(byVal lngPos, byRef ary)
'Sub WriteFooterTable
'Sub WritePageHeader
'Function LoadFilter
'Sub WriteItemFilter()

'**********************************************************
'*	Page Level variables
'**********************************************************

	Dim maryCustomValues
	Dim mblnAutoShowTable
	Dim mblnShowDetail
	Dim mblnShowFilter
	Dim mblnShowHeader
	Dim mblnShowSummary
	Dim mbytSummaryTableHeight
	Dim mclsItem
	Dim mlngAbsolutePage
	Dim mlngMaxRecords
	Dim mlngPageCount
	Dim mradTextSearch
	Dim mstrAction
	Dim mstrItemTitle
	Dim mstrShow
	Dim mstrSortField
	Dim mstrSortOrder
	Dim mstrsqlWhere
	Dim mstrTextSearch
	Dim mvntID

'**********************************************************
'*	Begin Page Code
'**********************************************************

	mstrPageTitle = "Product Inventory Administration"
	mstrssAddonVersion = "2.00.001"

	mlngMaxRecords = LoadRequestValue("PageSize")
	If len(mlngMaxRecords) = 0 Then mlngMaxRecords = 10

	mblnShowHeader = True
	mblnShowDetail = False

	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	mstrAction = LoadRequestValue("Action")
	If Len(mstrAction) = 0 Then mstrAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	
	mstrSortField = LoadRequestValue("SortField")
	mstrSortOrder = LoadRequestValue("SortOrder")

    Set mclsItem = New clsItem
    With mclsItem

		Select Case mstrAction
			Case "New", "Update"
				.Update
			Case "Delete"
				.Delete mvntID
				mvntID = ""
			Case "viewItem"
				mvntID = LoadRequestValue("ViewID")
			Case "Filter"
				mvntID = ""
		End Select
	    
		If .Load Then 

		End If
	
		Call WriteHeader("body_onload();",True)
%>
<script LANGUAGE=javascript>
<!--

var theDataForm;
var strDetailTitle = "<%= mstrItemTitle %>";
var blnIsDirty;
var strSubSection = "Status";

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
	DisplayMainSection('Summary');
}

function DisplaySection(strSection)
{
var arySections = new Array('General');

	frmData.Show.value = strSection;

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
}

function DisplayMainSection(strSection)
{

	var arySections = new Array('Filter', 'Summary');

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


function btnNewItem_onclick(theButton)
{
var theForm = theButton.form;

	SetDefaults(theForm);
    document.all("spanprodName").innerHTML = theDataForm.btnUpdate.value;

}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete this?");
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

function SetDefaults(theForm)
{
<%  
Dim i

If isArray(maryCustomValues) Then 
	For i = 0 To UBound(maryCustomValues)
		Response.Write "theForm." & maryCustomValues(i)(enCustomField_FieldName) & ".value = " & Chr(34) & Chr(34) & ";" & vbcrlf
	Next 'i
End If
%>
    
    
return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.SortField.value = strColumn;
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

function viewItem(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "viewItem";
	theDataForm.submit();
	return false;
}

function ValidInput(theForm)
{
var  strSection = frmData.Show.value;

	theDataForm.submit();
    return(true);
}

//-->
</script>
<%
End With

Call WriteFormOpener
Response.Write mclsItem.OutputMessage
%>
<ul style="MARGIN-LEFT: 18pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;">
<li><a href="ssInventoryImportFile.asp">Load Inventory File</a></li>
</ul>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplayMainSection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your filter criteria here.">&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View items which meet the specified filter criteria">&nbsp;Summaries&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" id="btnHelp" name="btnHelp" title="Release Version <%= mstrssAddonVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
	<%
		Call WriteItemFilter
		Dim pobjTmpOrders
		Set pobjTmpOrders = mclsItem.Records
	%>
	<table class='tbl' width='100%' style="border-collapse: collapse" cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblSummary' rules="none">
	  <tr class='tblhdr'>
		<th colspan="2">Product</th>
		<th>Attribute</th>
		<th>Quantity</th>
	  </tr>
	  <%
	  Dim PrevinvenId
	  Dim PrevCustID
	  Dim mblnNewVisitor
	  Dim mblnEvenProduct
	  Dim mstrClass
	  
	  mblnEvenProduct = False
	  
	  Do While Not pobjTmpOrders.EOF
		mblnNewVisitor = CBool(pobjTmpOrders.Fields("invenId").Value <> PrevinvenId)
		If mblnNewVisitor Then
			PrevinvenId = pobjTmpOrders.Fields("invenId").Value
			mblnEvenProduct = Not mblnEvenProduct
			
			If mblnEvenProduct Then
				mstrClass = " class=""Inactive"""
			Else
				mstrClass = ""
			End If

		%>
		<tr <%= mstrClass %>>
			<td><%= pobjTmpOrders.Fields("invenProdId").Value %></td>
			<td><%= pobjTmpOrders.Fields("prodName").Value %></td>
			<td><%= pobjTmpOrders.Fields("invenAttName").Value %></td>
			<td><%= pobjTmpOrders.Fields("invenInStock").Value %></td>
		</tr>
		<%
		End If	'mblnNewVisitor
		
		pobjTmpOrders.MoveNext
	  Loop
	  %>
	</table>
	<%	
		Call ReleaseObject(pobjTmpOrders)

	%>
	</td>
  </tr>
</table>
</FORM>
</BODY>
</HTML>
<%

	Call ReleaseObject(cnn)
    If Response.Buffer Then Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Sub WriteFormOpener
%>
<form id="frmData" name="frmData" onsubmit="return ValidInput(this);" method="post" action="ssInventoryList.asp">
<input type=hidden id="ViewID" name="ViewID">
<input type=hidden id=Action name=Action value="Update">
<input type=hidden id=blnShowSummary name=blnShowSummary value="">
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=Show name=Show value="<%= mstrShow %>">
<input type=hidden id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type=hidden id="SortField" name="SortField" value="<%= mstrSortField %>">
<input type=hidden id="SortOrder" name="SortOrder" value="<%= mstrSortOrder %>">

<% End Sub	'WriteFormOpener %>

<%
'**************************************************************************************************************************************************

Sub WriteFooterTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
        <input class='butn' title='Create a new Item' id=btnNewItem name=btnNewItem type=button value='New' onclick='return btnNewItem_onclick(this)'>&nbsp;
        <input class='butn' title="Reset" id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)' disabled>&nbsp;&nbsp;
        <input class='butn' title="Delete this Item" id=btnDelete name=btnDelete type=button value='Delete' onclick='return btnDelete_onclick(this)'>
        <input class='butn' title="Save changes" id=btnUpdate name=btnUpdate type=button value='Save Changes' onclick='return ValidInput(this.form);'>
    </td>
  </tr>
</table>
<%
End Sub	'WriteFooterTable

'************************************************************************************************************************************

Sub WritePageHeader
%>
<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
    <th>&nbsp;</th>
    <th align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br />
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>
	</th>
  </tr>
</table>
<% 
End Sub	'WritePageHeader 

'***********************************************************************************************

Function LoadFilter

Dim pstrSelFilter
Dim pstrsqlWhere

	'modified so could link in directly
	mradTextSearch = LoadRequestValue("radTextSearch")
	mstrTextSearch = trim(LoadRequestValue("TextSearch"))
	If (Len(mradTextSearch) > 0) And (Len(mstrTextSearch) > 0) Then
		pstrsqlWhere =  maryCustomValues(mradTextSearch)(enCustomField_FieldName) & " Like '%" & sqlSafe(mstrTextSearch) & "%'"
	End If

	For i = 0 To UBound(maryCustomValues)
		If maryCustomValues(i)(enCustomField_DisplayType) = enDisplayType_select Then
			pstrSelFilter = Trim(Request.Form("selFilter" & i ))
			If len(pstrSelFilter) > 0 then
				If len(pstrsqlWhere) > 0 Then
					pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					'pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
				Else
					pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					'pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
				End If
			End If
		End If
	Next 'i

	LoadFilter = pstrsqlWhere
	
End Function    'LoadFilter

'******************************************************************************************************************************************************************

Sub WriteItemFilter()

Dim i
Dim plngradTextCounter: plngradTextCounter = 0
Dim plng
%>
<script LANGUAGE=javascript>
<!--

function btnFilter_onclick(theButton)
{
var theForm = theButton.form;

  theForm.Action.value = "Filter";
  theForm.submit();
  return(true);
}

//-->
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
    <td valign="top">
        Filter on<br />
		<%
		
		%>
        <input type="radio" value="" <% if mradTextSearch="" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch"><label for="radTextSearch">Do Not Include</label>
        <br />containing the text<br />
        <input type=enDisplayType_textbox name="TextSearch" size="20" value="<%= EncodeString(mstrTextSearch,True) %>" ID="TextSearch">
	</td>
	
	<td valign="top" align="center">
		<%
		Dim pstrSelFilter
		%>

	</td>
	<td>
	  <input class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"><br />
	</td>
  </tr>
</table>
<% End Sub	'WriteItemFilter %>