<%Option Explicit
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version:	2.00.001		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		August 18, 2003											*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clssfLocalesState
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError

'database variables
Private pstrloclstAbbreviation
Private pbytloclstLocaleIsActive
Private pstrloclstName
Private pdblloclstTax
Private pbytloclstTaxIsActive

Private pstrloclstAbbreviation2

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
    Set pRS = Nothing
End Sub

'***********************************************************************************************

Public Property Let Recordset(oRS)
    set pRS = oRS
End Property

Public Property Get Recordset()
    set Recordset = pRS
End Property


Public Property Get Message()
    Message = pstrMessage
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


Public Property Get loclstAbbreviation()
    loclstAbbreviation = pstrloclstAbbreviation
End Property

Public Property Get loclstLocaleIsActive()
    loclstLocaleIsActive = pbytloclstLocaleIsActive
End Property

Public Property Get loclstName()
    loclstName = pstrloclstName
End Property

Public Property Get loclstTax()
    loclstTax = pdblloclstTax
End Property

Public Property Get loclstTaxIsActive()
    loclstTaxIsActive = pbytloclstTaxIsActive
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    pstrloclstAbbreviation = trim(rs("loclstAbbreviation"))
    pbytloclstLocaleIsActive = trim(rs("loclstLocaleIsActive"))
    pstrloclstName = trim(rs("loclstName"))
    pdblloclstTax = trim(rs("loclstTax"))
    pbytloclstTaxIsActive = trim(rs("loclstTaxIsActive"))

End Sub 'LoadValues

Private Sub ClearValues

    pstrloclstAbbreviation = ""
    pstrloclstName = ""
    pdblloclstTax = ""
    pbytloclstLocaleIsActive = False
    pbytloclstTaxIsActive = False

End Sub 'ClearValues

Private Sub LoadFromRequest

    With Request.Form
        pstrloclstAbbreviation2 = Trim(.Item("loclstAbbreviation2"))
        pstrloclstAbbreviation = Trim(.Item("loclstAbbreviation"))
        pstrloclstName = Trim(.Item("loclstName"))
        pdblloclstTax = Trim(.Item("loclstTax"))
        pbytloclstTaxIsActive = cBool(lCase(.Item("loclstTaxIsActive") = "on")) * -1
        pbytloclstLocaleIsActive = cBool(lCase(.Item("loclstLocaleIsActive") = "on")) * -1
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "loclstAbbreviation='" & lngID & "'"
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues (pRS)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function LoadAll(sqlWhere)

'On Error Resume Next

    Set pRS = GetRS("Select * from sfLocalesState " & sqlWhere)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(strloclstAbbreviation)

Dim sql

'On Error Resume Next

    sql = "Delete from sfLocalesState where loclstAbbreviation = '" & strloclstAbbreviation & "'"
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Activate(lngID,blnActivate)

Dim sql

On Error Resume Next

	if blnActivate then
		sql = "Update sfLocalesState Set loclstLocaleIsActive=1 where loclstAbbreviation='" & lngID & "'"
    else
		sql = "Update sfLocalesState Set loclstLocaleIsActive=0 where loclstAbbreviation='" & lngID & "'"
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "State successfully updated."
        Activate = True
    Else
        pstrMessage = Err.Description
        Activate = False
    End If

End Function    'Activate

'***********************************************************************************************

Public Function ActivateTax(lngID,blnActivate)

Dim sql

On Error Resume Next

	if blnActivate then
		sql = "Update sfLocalesState Set loclstTaxIsActive=1 where loclstAbbreviation='" & lngID & "'"
    else
		sql = "Update sfLocalesState Set loclstTaxIsActive=0 where loclstAbbreviation='" & lngID & "'"
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "State successfully updated."
        ActivateTax = True
    Else
        pstrMessage = Err.Description
        ActivateTax = False
    End If

End Function    'ActivateTax

'***********************************************************************************************

Public Function Update()

Dim sql
Dim rs
Dim strErrorMessage
Dim blnAdd

On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
        sql = "Select * from sfLocalesState where loclstAbbreviation = '" & pstrloclstAbbreviation2 & "'"
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

		rs("loclstAbbreviation") = pstrloclstAbbreviation
        rs("loclstName") = pstrloclstName
        rs("loclstTax") = pdblloclstTax
        rs("loclstTaxIsActive") = pbytloclstTaxIsActive
        rs("loclstLocaleIsActive") = pbytloclstLocaleIsActive

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The State abbreviation you entered (" & pstrloclstAbbreviation & ") is already in use.<br />Please enter a different abbreviation.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
        
        pstrloclstAbbreviation = rs("loclstAbbreviation")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrloclstName & " was successfully added."
            Else
                pstrMessage = pstrloclstName & " was successfully updated."
            End If
        Else
            pblnError = True
        End If
    Else
        pblnError = True
    End If

    Update = (not pblnError)
    
    Application.Contents.Remove("StateArray")

End Function    'Update

'***********************************************************************************************


Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim aSortHeader(5,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "  <tr class='tblhdr'>"

	If (len(strradSortOrder) = 0) or (strradSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort States in descending order"
		aSortHeader(2,0) = "Sort State Abbreviations in descending order"
		aSortHeader(3,0) = "Sort Tax Rates in descending order"
		aSortHeader(4,0) = "Sort active taxes in descending order"
		aSortHeader(5,0) = "Sort active states in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort States in ascending order"
		aSortHeader(2,0) = "Sort State Abbreviations in ascending order"
		aSortHeader(3,0) = "Sort Tax Rates in ascending order"
		aSortHeader(4,0) = "Sort active taxes in ascending order"
		aSortHeader(5,0) = "Sort active states in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;State"
	aSortHeader(2,1) = "Abbr"
	aSortHeader(3,1) = "Tax Rate"
	aSortHeader(4,1) = "Tax Active"
	aSortHeader(5,1) = "Is Active"

	if len(strradOrderBy) > 0 Then
		pstrOrderBy = strradOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 1 to 5
		If cInt(pstrOrderBy) = i Then
			If (pstrSortOrder = "ASC") Then
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
			Else
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
			End If
		Else
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
		End If
	next 'i

	Response.Write "<tr><td colspan=5>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' " _
				 & ">"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
				  
    If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrAbbr = Trim(prs("loclstAbbreviation"))
 			pstrTitle = "Click to edit " & prs("loclstName") & "."
			pstrURL = "ShippingAdmin.asp?Action=View&ID=" & pstrAbbr

			if pstrAbbr = pstrloclstAbbreviation then
        		Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
			else
				if cBool(pRS("loclstLocaleIsActive")) then
	                Response.Write "<TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				else
	                Response.Write "<TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
        		end if
        	end if

            .Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
									" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
									" onclick=" & chr(34) & "ViewDetail('" & pstrAbbr & "');" & chr(34) & ">" & prs("loclstName") & "</a></TD>" & vbCrLf
        	.Write "  <TD>" & pRS("loclstAbbreviation") & "</TD>"
        	.Write "  <TD>" & pRS("loclstTax") & "&nbsp;</TD>"
			if cBool(pRS("loclstTaxIsActive")) then
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateTax('" & pstrAbbr & "',false);" & chr(34) & " title='Click to make tax for " & prs("loclstName") & " inactive.'>Active</a></TD>" & vbCrLf
			else
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateTax('" & pstrAbbr & "',true);" & chr(34) & " title='Click to make tax for " & prs("loclstName") & " active.'>Inactive</a></TD>" & vbCrLf
        	end if
			if cBool(pRS("loclstLocaleIsActive")) then
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateLocale('" & pstrAbbr & "',false);" & chr(34) & " title='Click to make " & prs("loclstName") & " inactive.'>Active</a></TD>" & vbCrLf
			else
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateLocale('" & pstrAbbr & "',true);" & chr(34) & " title='Click to make " & prs("loclstName") & " active.'>Inactive</a></TD>" & vbCrLf
        	end if

            prs.MoveNext
        Next
    Else
        .Write "<TR><TD colspan=5 align=center><h3>There are no States</h3></TD></TR>"
		Call ClearValues
    End If
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
    
    End With

End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If len(pstrloclstName) = 0 <> 0 Then
        strError = strError & "Please enter a State." & cstrDelimeter
    End If

    If len(pstrloclstAbbreviation) = 0 <> 0 Then
        strError = strError & "Please enter a State abbreviation." & cstrDelimeter
    End If

    If Not IsNumeric(pdblloclstTax) And Len(pdblloclstTax) <> 0 Then
        strError = strError & "Please enter a number for the State tax." & cstrDelimeter
    ElseIf Len(pdblloclstTax) = 0 Then
        strError = strError & "Please enter a numeric value for the State tax." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)

End Function 'ValidateValues

End Class   'clssfLocalesState

Sub LoadFilter

dim pstrsqlWhere
dim pstrOrderBy

	strradFilter = Request.Form("radFilter")
	strradFilter2 = Request.Form("radFilter2")
	strradOrderBy = Request.Form("OrderBy")
	strradSortOrder = Request.Form("SortOrder")
	blnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	blnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")

'Build Filter

	Select Case strradFilter	'Show Countries that are
		Case "0"	'All
		Case "1"	'Active
			pstrsqlWhere = "loclstLocaleIsActive=1"
		Case "2"	'Inactive
			pstrsqlWhere = "loclstLocaleIsActive=0"
	End Select	
	
	Select Case strradFilter2	'Show TaxRates that are
		Case "0"	'All
		Case "1"	'Active
			If len(pstrsqlWhere) = 0 Then
				pstrsqlWhere = "loclstTaxIsActive=1"
			Else
				pstrsqlWhere = pstrsqlWhere & " and loclstTaxIsActive=1"
			End If
		Case "2"	'Inactive
			If len(pstrsqlWhere) = 0 Then
				pstrsqlWhere = "loclstTaxIsActive=0"
			Else
				pstrsqlWhere = pstrsqlWhere & " and loclstTaxIsActive=0"
			End If
	End Select	

	If len(strradOrderBy) = 0 Then strradOrderBy = 1
	Select Case strradOrderBy	'Order By
		Case "1"	'State
			pstrOrderBy = "loclstName"
		Case "2"	'State Abbr
			pstrOrderBy = "loclstAbbreviation"
		Case "3"	'Tax Rate
			pstrOrderBy = "loclstTax"
		Case "4"	'Tax Active
			pstrOrderBy = "loclstTaxIsActive"
		Case "5"	'State Active
			pstrOrderBy = "loclstLocaleIsActive"
	End Select	

	If len(pstrsqlWhere) > 0 then
		mstrsqlWhere = " Where " & pstrsqlWhere
	End If

	If len(pstrOrderBy) > 0 then
		mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & strradSortOrder
	Else
		mstrsqlWhere = mstrsqlWhere & " Order By loclstName"
	End If
	
End Sub    'LoadFilter

mstrPageTitle = "State Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfLocalesState
Dim strradFilter, strradFilter2, strradOrderBy, strradSortOrder, blnShowFilter, blnShowSummary
Dim mvntID
Dim mstrsqlWhere

	Call LoadFilter
    mAction = Request.Form("Action")
    mvntID = Request.Form("loclstAbbreviation")
       
    Set mclssfLocalesState = New clssfLocalesState
    
    Select Case mAction
        Case "New", "Update"
            mclssfLocalesState.Update
            If mclssfLocalesState.LoadAll(mstrsqlWhere) Then mclssfLocalesState.Find mvntID
        Case "Delete"
            mclssfLocalesState.Delete Request.Form("loclstAbbreviation")
            mclssfLocalesState.LoadAll mstrsqlWhere
        Case "ViewDetail"
            If mclssfLocalesState.LoadAll(mstrsqlWhere) Then mclssfLocalesState.Find mvntID
        Case "Activate", "Deactivate"
            mclssfLocalesState.Activate mvntID, (mAction= "Activate")
            If mclssfLocalesState.LoadAll(mstrsqlWhere) Then mclssfLocalesState.Find mvntID
        Case "ActivateTax", "DeactivateTax"
            mclssfLocalesState.ActivateTax mvntID, (mAction= "ActivateTax")
            If mclssfLocalesState.LoadAll(mstrsqlWhere) Then mclssfLocalesState.Find mvntID
        Case Else
            mclssfLocalesState.LoadAll mstrsqlWhere
    End Select
    
	Call WriteHeader("body_onload();",True)
    With mclssfLocalesState
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theKeyField;
var theDataForm;
var strDetailTitle = "<%= .loclstName %> Details";

function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = document.frmData.loclstAbbreviation;
<% If blnShowFilter then Response.Write "DisplayFilter();" & vbcrlf %>
<% If blnShowSummary then Response.Write "DisplaySummary();" & vbcrlf %>
}

function SetDefaults(theForm)
{
    theForm.loclstAbbreviation.value = "";
    theForm.loclstAbbreviation2.value = "";
    theForm.loclstName.value = "";
    theForm.loclstTax.value = "0";
    theForm.loclstLocaleIsActive.checked = false;
    theForm.loclstTaxIsActive.checked = false;
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add State";
    theForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.loclstName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }
    {
    return(false);
    }
}

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function ValidInput(theButton)
{
var theForm = theButton.form;

  if (isEmpty(theForm.loclstName,"Please enter a State name.")) {return(false);}
  if (isEmpty(theForm.loclstAbbreviation,"Please enter a State abbreviation.")) {return(false);}
  if (!isNumeric(theForm.loclstTax,true,"Please enter a numeric value for the State tax.")) {return(false);}
  theForm.Action.value = "Update";
  theForm.submit();
}

function ActivateTax(theValue,bln)
{
	theKeyField.value = theValue;
	if (bln)
	{
	theDataForm.Action.value = "ActivateTax";
	theDataForm.submit();
	return false;
	}
	{
	theDataForm.Action.value = "DeactivateTax";
	theDataForm.submit();
	return false;
	}
}

function ActivateLocale(theValue,bln)
{
	theKeyField.value = theValue;
	if (bln)
	{
	theDataForm.Action.value = "Activate";
	theDataForm.submit();
	return false;
	}
	{
	theDataForm.Action.value = "Deactivate";
	theDataForm.submit();
	return false;
	}
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theDataForm.Action.value = "ViewDetail";
	theDataForm.submit();
	return false;
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

//-->
</SCRIPT>

<BODY onload="body_onload();">
<CENTER>
<TABLE border=0 cellPadding=5 cellSpacing=1 width="95%">
  <TR>
    <TH><div class="pagetitle "><%= mstrPageTitle %></div></TH>
    <TH>&nbsp;</TH>
    <TH align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br />
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>

	</TH>
  </TR>
</TABLE>
<%= .OutputMessage %>

<FORM action='sfLocalesStateAdmin.asp' id=frmData name=frmData method=post>
<input type=hidden id=loclstAbbreviation2 name=loclstAbbreviation2 value=<%= .loclstAbbreviation %>>
<input type=hidden id=Action name=Action value=''>
<input type=hidden id=blnShowFilter name=blnShowFilter value=''>
<input type=hidden id=blnShowSummary name=blnShowSummary value=''>
<input type=hidden id=OrderBy name=OrderBy value='<%= strradOrderBy %>'>
<input type=hidden id=SortOrder name=SortOrder value='<%= strradSortOrder %>'>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
	<th>Show Countries that are</th>
	<th>Show Tax Rates that are</th>
	<th>&nbsp;</th>
  </tr>
  <TR>
    <TD valign="top">
        &nbsp;&nbsp;<input type="radio" value="1" <% if strradFilter="1" then Response.Write "Checked" %> name="radFilter">Active<br />
        &nbsp;&nbsp;<input type="radio" value="2" <% if strradFilter="2" then Response.Write "Checked" %> name="radFilter">Inactive<br />
        &nbsp;&nbsp;<input type="radio" value="0" <% if (strradFilter="0" or strradFilter="") then Response.Write "Checked" %> name="radFilter">All
	</TD>
    <TD valign="top">
        &nbsp;&nbsp;<input type="radio" value="1" <% if strradFilter2="1" then Response.Write "Checked" %> name="radFilter2">Active<br />
        &nbsp;&nbsp;<input type="radio" value="2" <% if strradFilter2="2" then Response.Write "Checked" %> name="radFilter2">Inactive<br />
        &nbsp;&nbsp;<input type="radio" value="0" <% if (strradFilter2="0" or strradFilter="") then Response.Write "Checked" %> name="radFilter2">All
	</TD>
    <TD valign=middle align=center>
		<INPUT class='butn' id=btnSubmit name=btnSubmit type=submit value="Apply Filter">
	</TD>
</tr>
</TABLE>

<%= .OutputSummary %>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="left">
<colgroup align="left">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="3" align=center><span id="spanDetailTitle"><%= .loclstName %> Details</span></th>
  </tr>
      <TR>
        <TD class="label">&nbsp;<LABEL id=lblloclstName for=loclstName>State:</LABEL></TD>
        <TD><INPUT id=loclstName name=loclstName Value='<%= .loclstName %>' maxlength=50 size=50></TD>
        <TD><INPUT type=checkbox id=loclstLocaleIsActive name=loclstLocaleIsActive <% If .loclstLocaleIsActive Then Response.Write "Checked" %>><LABEL id=lblloclstLocaleIsActive for=loclstLocaleIsActive>State is active</LABEL></TD>
      </TR>
      <TR>
        <TD class="label">&nbsp;<LABEL id=lblloclstAbbreviation for=loclstAbbreviation>State Abbreviation:</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=loclstAbbreviation name=loclstAbbreviation Value='<%= .loclstAbbreviation %>' maxlength=3 size=3></TD>
      </TR>
      <TR>
        <TD class="label">&nbsp;<LABEL id=lblloclstTax for=loclstTax>State Tax:</LABEL></TD>
        <TD><INPUT id=loclstTax name=loclstTax Value='<%= .loclstTax %>'></TD>
        <TD><INPUT type=checkbox id=loclstTaxIsActive name=loclstTaxIsActive <% If .loclstTaxIsActive Then Response.Write "Checked" %>>&nbsp;<LABEL id=lblloclstTaxIsActive for=loclstTaxIsActive>State Tax is active</LABEL></TD>
      </TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick(this)'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=button value='Save Changes' onclick='ValidInput(this);'>
    </TD>
    <TD>&nbsp;</TD>
  </TR>
</TABLE>
</FORM>

</CENTER>
</BODY>
</HTML>
<%
    End With
    Set mclssfLocalesState = Nothing
    Set cnn = Nothing
    Response.Flush
%>
