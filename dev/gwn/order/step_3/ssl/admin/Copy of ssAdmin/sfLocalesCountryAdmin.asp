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

Class clssfLocalesCountry
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError

'database variables
Private pstrloclctryAbbreviation
Private pbytloclctryLocalIsActive
Private pstrloclctryName
Private pdblloclctryTax
Private pbytloclctryTaxIsActive
Private pbytloclctryFraudRating

Private pstrloclctryAbbreviation2

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


Public Property Get loclctryAbbreviation()
    loclctryAbbreviation = pstrloclctryAbbreviation
End Property

Public Property Get loclctryLocalIsActive()
    loclctryLocalIsActive = pbytloclctryLocalIsActive
End Property

Public Property Get loclctryName()
    loclctryName = pstrloclctryName
End Property

Public Property Get loclctryTax()
    loclctryTax = pdblloclctryTax
End Property

Public Property Get loclctryTaxIsActive()
    loclctryTaxIsActive = pbytloclctryTaxIsActive
End Property

Public Property Get loclctryFraudRating()
    loclctryFraudRating = pbytloclctryFraudRating
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    pstrloclctryAbbreviation = trim(rs("loclctryAbbreviation"))
    pbytloclctryLocalIsActive = trim(rs("loclctryLocalIsActive"))
    pstrloclctryName = trim(rs("loclctryName"))
    pdblloclctryTax = trim(rs("loclctryTax"))
    pbytloclctryTaxIsActive = trim(rs("loclctryTaxIsActive"))
    pbytloclctryFraudRating = trim(rs("loclctryFraudRating"))

End Sub 'LoadValues

Private Sub ClearValues

    pstrloclctryAbbreviation = ""
    pstrloclctryName = ""
    pdblloclctryTax = ""
    pbytloclctryLocalIsActive = False
    pbytloclctryTaxIsActive = False
    pbytloclctryTaxIsActive = 0

End Sub 'ClearValues

Private Sub LoadFromRequest

    With Request.Form
        pstrloclctryAbbreviation2 = Trim(.Item("loclctryAbbreviation2"))
        pstrloclctryAbbreviation = Trim(.Item("loclctryAbbreviation"))
        pstrloclctryName = Trim(.Item("loclctryName"))
        pdblloclctryTax = Trim(.Item("loclctryTax"))
        pbytloclctryTaxIsActive = cBool(lCase(.Item("loclctryTaxIsActive") = "on")) * -1
        pbytloclctryFraudRating = Trim(.Item("loclctryFraudRating"))
        pbytloclctryLocalIsActive = cBool(lCase(.Item("loclctryLocalIsActive") = "on")) * -1
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "loclctryAbbreviation='" & lngID & "'"
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

    Set pRS = GetRS("Select * from sfLocalesCountry " & sqlWhere)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(strloclctryAbbreviation)

Dim sql

'On Error Resume Next

    sql = "Delete from sfLocalesCountry where loclctryAbbreviation = '" & strloclctryAbbreviation & "'"
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
		sql = "Update sfLocalesCountry Set loclctryLocalIsActive=1 where loclctryAbbreviation='" & lngID & "'"
    else
		sql = "Update sfLocalesCountry Set loclctryLocalIsActive=0 where loclctryAbbreviation='" & lngID & "'"
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "Country successfully updated."
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
		sql = "Update sfLocalesCountry Set loclctryTaxIsActive=1 where loclctryAbbreviation='" & lngID & "'"
    else
		sql = "Update sfLocalesCountry Set loclctryTaxIsActive=0 where loclctryAbbreviation='" & lngID & "'"
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "Country successfully updated."
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
        If Len(pstrloclctryAbbreviation) = 0 Then pstrloclctryAbbreviation = "''"

        sql = "Select * from sfLocalesCountry where loclctryAbbreviation = '" & pstrloclctryAbbreviation2 & "'"
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

		rs("loclctryAbbreviation") = pstrloclctryAbbreviation
        rs("loclctryName") = pstrloclctryName
        rs("loclctryTax") = pdblloclctryTax
        rs("loclctryTaxIsActive") = pbytloclctryTaxIsActive
        rs("loclctryFraudRating") = pbytloclctryFraudRating
        rs("loclctryLocalIsActive") = pbytloclctryLocalIsActive

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The country abbreviation you entered (" & pstrloclctryAbbreviation & ") is already in use.<BR>Please enter a different abbreviation.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        pstrloclctryAbbreviation = rs("loclctryAbbreviation")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrloclctryName & " was successfully added."
            Else
                pstrMessage = pstrloclctryName & " was successfully updated."
            End If
        Else
            pblnError = True
        End If
    Else
        pblnError = True
    End If

    Update = (not pblnError)

    Application.Contents.Remove("CountryArray")

End Function    'Update

'***********************************************************************************************


Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim aSortHeader(6,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "  <tr class='tblhdr'>"

	If (len(strradSortOrder) = 0) or (strradSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort Countries in descending order"
		aSortHeader(2,0) = "Sort Country Abbreviations in descending order"
		aSortHeader(3,0) = "Sort Tax Rates in descending order"
		aSortHeader(4,0) = "Sort active taxes in descending order"
		aSortHeader(5,0) = "Sort active countries in descending order"
		aSortHeader(6,0) = "Sort fraud score in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort Countries in ascending order"
		aSortHeader(2,0) = "Sort Country Abbreviations in ascending order"
		aSortHeader(3,0) = "Sort Tax Rates in ascending order"
		aSortHeader(4,0) = "Sort active taxes in ascending order"
		aSortHeader(5,0) = "Sort active countries in ascending order"
		aSortHeader(6,0) = "Sort fraud score in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;Country"
	aSortHeader(2,1) = "Country Abbreviation"
	aSortHeader(3,1) = "Tax Rate"
	aSortHeader(4,1) = "Tax Active"
	aSortHeader(5,1) = "Is Active"
	aSortHeader(6,1) = "Fraud"

	if len(strradOrderBy) > 0 Then
		pstrOrderBy = strradOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 1 to 6
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

	Response.Write "<tr><td colspan=6>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' " _
				 & ">"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "<COLGROUP align='center' width='10%'>"

    If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrAbbr = Trim(prs("loclctryAbbreviation"))
 			pstrTitle = "Click to edit " & prs("loclctryName") & "."
			pstrURL = "sfLocalesCountryAdmin.asp?Action=View&loclctryAbbreviation=" & pstrAbbr

			if pstrAbbr = pstrloclctryAbbreviation then
        		Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
			else
				if cBool(pRS("loclctryLocalIsActive")) then
	                Response.Write "<TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				else
	                Response.Write "<TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
        		end if
        	end if

            .Write "  <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=" & chr(34) & "#" & chr(34) & _
									" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
									" onclick=" & chr(34) & "ViewDetail('" & pstrAbbr & "');" & chr(34) & ">" & prs("loclctryName") & "</a></TD>" & vbCrLf
        	.Write "  <TD>" & pstrAbbr & "</TD>"
        	.Write "  <TD>" & pRS("loclctryTax") & "&nbsp;</TD>"
			if cBool(pRS("loclctryTaxIsActive")) then
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateTax('" & pstrAbbr & "',false);" & chr(34) & " title='Click to make tax for " & prs("loclctryName") & " inactive.'>Active</a></TD>" & vbCrLf
			else
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateTax('" & pstrAbbr & "',true);" & chr(34) & " title='Click to make tax for " & prs("loclctryName") & " active.'>Inactive</a></TD>" & vbCrLf
        	end if
			if cBool(pRS("loclctryLocalIsActive")) then
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateLocale('" & pstrAbbr & "',false);" & chr(34) & " title='Click to make " & prs("loclctryName") & " inactive.'>Active</a></TD>" & vbCrLf
			else
				.Write "  <TD><a href=" & chr(34) & "#" & chr(34) & _
										" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();'" & _
										" onclick=" & chr(34) & "ActivateLocale('" & pstrAbbr & "',true);" & chr(34) & " title='Click to make " & prs("loclctryName") & " active.'>Inactive</a></TD>" & vbCrLf
        	end if
			.Write "  <TD>" & pRS("loclctryFraudRating") & "</TD>"
            prs.MoveNext
        Next
    Else
        .Write "<TR><TD colspan=6 align=center><h3>There are no Countries</h3></TD></TR>"
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

    If len(pstrloclctryName) = 0 <> 0 Then
        strError = strError & "Please enter a country." & cstrDelimeter
    End If

    If len(pstrloclctryAbbreviation) = 0 <> 0 Then
        strError = strError & "Please enter a country abbreviation." & cstrDelimeter
    End If

    If Not IsNumeric(pdblloclctryTax) And Len(pdblloclctryTax) <> 0 Then
        strError = strError & "Please enter a number for the country tax." & cstrDelimeter
    ElseIf Len(pdblloclctryTax) = 0 Then
        strError = strError & "Please enter a numeric value for the country tax." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)

End Function 'ValidateValues

End Class   'clssfLocalesCountry

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
			pstrsqlWhere = "loclctryLocalIsActive=1"
		Case "2"	'Inactive
			pstrsqlWhere = "loclctryLocalIsActive=0"
	End Select	
	
	Select Case strradFilter2	'Show TaxRates that are
		Case "0"	'All
		Case "1"	'Active
			If len(pstrsqlWhere) = 0 Then
				pstrsqlWhere = "loclctryTaxIsActive=1"
			Else
				pstrsqlWhere = pstrsqlWhere & " and loclctryTaxIsActive=1"
			End If
		Case "2"	'Inactive
			If len(pstrsqlWhere) = 0 Then
				pstrsqlWhere = "loclctryTaxIsActive=0"
			Else
				pstrsqlWhere = pstrsqlWhere & " and loclctryTaxIsActive=0"
			End If
	End Select	

	If len(strradOrderBy) = 0 Then strradOrderBy = 1
	Select Case strradOrderBy	'Order By
		Case "1"	'Country
			pstrOrderBy = "loclctryName"
		Case "2"	'Country Abbr
			pstrOrderBy = "loclctryAbbreviation"
		Case "3"	'Tax Rate
			pstrOrderBy = "loclctryTax"
		Case "4"	'Tax Active
			pstrOrderBy = "loclctryTaxIsActive"
		Case "5"	'Country Active
			pstrOrderBy = "loclctryLocalIsActive"
		Case "6"	'Country Active
			pstrOrderBy = "loclctryFraudRating"
	End Select	

	If len(pstrsqlWhere) > 0 then
		mstrsqlWhere = " Where " & pstrsqlWhere
	End If

	If len(pstrOrderBy) > 0 then
		mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & strradSortOrder
	Else
		mstrsqlWhere = mstrsqlWhere & " Order By loclctryName"
	End If
	
End Sub    'LoadFilter

mstrPageTitle = "Country Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfLocalesCountry
Dim strradFilter, strradFilter2, strradOrderBy, strradSortOrder, blnShowFilter, blnShowSummary
Dim mvntID
Dim mstrsqlWhere

	Call LoadFilter
    mAction = Request.Form("Action")
    mvntID = Request.Form("loclctryAbbreviation")
       
    Set mclssfLocalesCountry = New clssfLocalesCountry
    
    Select Case mAction
        Case "New", "Update"
            mclssfLocalesCountry.Update
            If mclssfLocalesCountry.LoadAll(mstrsqlWhere) Then mclssfLocalesCountry.Find mvntID
        Case "Delete"
            mclssfLocalesCountry.Delete Request.Form("loclctryAbbreviation")
            mclssfLocalesCountry.LoadAll mstrsqlWhere
        Case "ViewDetail"
            If mclssfLocalesCountry.LoadAll(mstrsqlWhere) Then mclssfLocalesCountry.Find mvntID
        Case "Activate", "Deactivate"
            mclssfLocalesCountry.Activate mvntID, (mAction= "Activate")
            If mclssfLocalesCountry.LoadAll(mstrsqlWhere) Then mclssfLocalesCountry.Find mvntID
        Case "ActivateTax", "DeactivateTax"
            mclssfLocalesCountry.ActivateTax mvntID, (mAction= "ActivateTax")
            If mclssfLocalesCountry.LoadAll(mstrsqlWhere) Then mclssfLocalesCountry.Find mvntID
        Case Else
            mclssfLocalesCountry.LoadAll mstrsqlWhere
    End Select
    
	Call WriteHeader("body_onload();",True)
    With mclssfLocalesCountry
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theKeyField;
var theDataForm;
var strDetailTitle = "<%= .loclctryName %> Details";


function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = document.frmData.loclctryAbbreviation;
<% If blnShowFilter then Response.Write "DisplayFilter();" & vbcrlf %>
<% If blnShowSummary then Response.Write "DisplaySummary();" & vbcrlf %>
}

function SetDefaults(theForm)
{
    theForm.loclctryAbbreviation.value = "";
    theForm.loclctryAbbreviation2.value = "";
    theForm.loclctryName.value = "";
    theForm.loclctryTax.value = "0";
    theForm.loclctryLocalIsActive.checked = false;
    theForm.loclctryTaxIsActive.checked = false;
    theForm.loclctryTaxIsActive.value = "0";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add Country";
    theForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.loclctryName.value + "?");
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

  if (isEmpty(theForm.loclctryName,"Please enter a country name.")) {return(false);}
  if (isEmpty(theForm.loclctryAbbreviation,"Please enter a country abbreviation.")) {return(false);}
  if (!isNumeric(theForm.loclctryTax,true,"Please enter a numeric value for the country tax.")) {return(false);}
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
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br>
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>

	</TH>
  </TR>
</TABLE>
<%= .OutputMessage %>

<FORM action='sfLocalesCountryAdmin.asp' id=frmData name=frmData method=post>
<input type=hidden id=loclctryAbbreviation2 name=loclctryAbbreviation2 value=<%= .loclctryAbbreviation %>>
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
        &nbsp;&nbsp;<input type="radio" value="1" <% if strradFilter="1" then Response.Write "Checked" %> name="radFilter">Active<br>
        &nbsp;&nbsp;<input type="radio" value="2" <% if strradFilter="2" then Response.Write "Checked" %> name="radFilter">Inactive<br>
        &nbsp;&nbsp;<input type="radio" value="0" <% if (strradFilter="0" or strradFilter="") then Response.Write "Checked" %> name="radFilter">All
	</TD>
    <TD valign="top">
        &nbsp;&nbsp;<input type="radio" value="1" <% if strradFilter2="1" then Response.Write "Checked" %> name="radFilter2">Active<br>
        &nbsp;&nbsp;<input type="radio" value="2" <% if strradFilter2="2" then Response.Write "Checked" %> name="radFilter2">Inactive<br>
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
	<th colspan="3" align=center><span id="spanDetailTitle"><%= .loclctryName %> Details</span></th>
  </tr>
      <TR>
        <TD class="label"><LABEL id=lblloclctryName for=loclctryName>Country:</LABEL></TD>
        <TD>&nbsp;<INPUT id=loclctryName name=loclctryName Value='<%= .loclctryName %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=loclctryLocalIsActive name=loclctryLocalIsActive <% If .loclctryLocalIsActive Then Response.Write "Checked" %>><LABEL id=lblloclctryLocalIsActive for=loclctryLocalIsActive>Country is active</LABEL></TD>
      </TR>
      <TR>
        <TD class="label"><LABEL id=lblloclctryAbbreviation for=loclctryAbbreviation>Country Abbreviation:</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=loclctryAbbreviation name=loclctryAbbreviation Value='<%= .loclctryAbbreviation %>' maxlength=3 size=3></TD>
      </TR>
      <TR>
        <TD class="label"><LABEL id=lblloclctryTax for=loclctryTax>Country Tax:</LABEL></TD>
        <TD>&nbsp;<INPUT id=loclctryTax name=loclctryTax Value='<%= .loclctryTax %>'></TD>
        <TD>&nbsp;<INPUT type=checkbox id=loclctryTaxIsActive name=loclctryTaxIsActive <% If .loclctryTaxIsActive Then Response.Write "Checked" %>>&nbsp;<LABEL id=lblloclctryTaxIsActive for=loclctryTaxIsActive>Country Tax is active</LABEL></TD>
      </TR>
      <TR>
        <TD class="label"><LABEL id="lblloclctryFraudRating" for=loclctryFraudRating>Fraud Rating:</LABEL></TD>
        <TD>&nbsp;<INPUT id="loclctryFraudRating" name=loclctryFraudRating Value='<%= .loclctryFraudRating %>'></TD>
        <TD>&nbsp;</TD>
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
    Set mclssfLocalesCountry = Nothing
    Set cnn = Nothing
    Response.Flush
%>
