<% Option Explicit 
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

Class clsShipping
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private pstrACTIVE
Private pstrCODE
Private pstrEDIT
Private plngID
Private pstrMETHODS
Private pstrRATES

'***********************************************************************************************

Private Sub class_Initialize()
	cstrDelimeter = ";"
End Sub

Private Sub class_Terminate()
	On Error Resume Next
	set pRS = nothing
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


Public Property Get ACTIVE()
    ACTIVE = pstrACTIVE
End Property


Public Property Get CODE()
    CODE = pstrCODE
End Property


Public Property Get EDIT()
    EDIT = pstrEDIT
End Property


Public Property Get ID()
    ID = plngID
End Property


Public Property Get METHODS()
    METHODS = pstrMETHODS
End Property


Public Property Get RATES()
    RATES = pstrRATES
End Property


'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    with pRS
		if .RecordCount > 0 then
			.MoveFirst
			If len(lngID) <> 0 then
				.Find "shipID=" & lngID
			Else
				.MoveLast
			End If
			if not .EOF then LoadValues(pRS)
		end if
	end with
  
End Function    'Find

'***********************************************************************************************

Private Sub LoadValues(rs)

    pstrACTIVE = trim(rs("shipIsActive"))
    pstrCODE = trim(rs("shipCode"))
    pstrEDIT = trim(rs("shipEdit"))
    plngID = trim(rs("shipID"))
    pstrMETHODS = trim(rs("shipMethod"))
    pstrRATES = trim(rs("shipRates"))

End Sub 'LoadValues

Public Sub LoadFromRequest

    With Request.Form
        pstrACTIVE = (UCase(.Item("ACTIVE")) = "ON")
        pstrCODE = Trim(.Item("CODE"))
        pstrEDIT = (UCase(.Item("EDIT")) = "ON")
        plngID = Trim(.Item("ID"))
        pstrMETHODS = Trim(.Item("METHODS"))
        pstrRATES = Trim(.Item("RATES"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function LoadAll()

On Error Resume Next

    Set pRS = GetRS("Select * from sfShipping " & mstrsqlWhere)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(lngID)

Dim sql
Dim rs

On Error Resume Next

    sql = "Delete from sfShipping where shipID = " & lngID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Shipping method successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Activate(lngID,blnActivate)

Dim sql

'On Error Resume Next

	if blnActivate then
		sql = "Update sfShipping Set shipIsActive='1' where shipID = " & lngID
    else
		sql = "Update sfShipping Set shipIsActive='0' where shipID = " & lngID
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully updated."
        Activate = True
    Else
        pstrMessage = Err.Description
        Activate = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Update()

Dim sql
Dim rs
Dim strErrorMessage
Dim blnAdd

'On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
        If Len(plngID) = 0 Then plngID = 0

        sql = "Select * from sfShipping where shipID = " & plngID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("shipIsActive") = pstrACTIVE * -1
        rs("shipCode") = pstrCODE
        rs("shipEdit") = pstrEDIT * -1
        rs("shipMethod") = pstrMETHODS
        rs("shipRates") = pstrRATES

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngID = rs("shipID")
        rs.Close
        Set rs = Nothing
        
        mlngID = plngID

        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The record was successfully added."
            Else
                pstrMessage = "The record was successfully updated."
            End If
        Else
            pblnError = True
        End If
    Else
		pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrRATES) = 0 Then
        strError = strError & "Please enter a value for the RATES." & cstrDelimeter
    End If
    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues

Sub OutputSummary

Dim i
Dim pstrTitle, pstrURL
Dim aSortHeader(4,1)
Dim pstrOrderBy
 
	With Response
	
    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"
    .Write "  <tr class='tblhdr'>"


	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort Shipping Methods in descending order"
		aSortHeader(2,0) = "Sort Shipping Codes in descending order"
		aSortHeader(3,0) = "Sort Shipping Rates in descending order"
		aSortHeader(4,0) = "Sort Active Shipping Methods in descending order"
	Else
		aSortHeader(1,0) = "Sort Shipping Methods in ascending order"
		aSortHeader(2,0) = "Sort Shipping Codes in ascending order"
		aSortHeader(3,0) = "Sort Shipping Rates in ascending order"
		aSortHeader(4,0) = "Sort Active Shipping Methods in ascending order"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;Shipping Method"
	aSortHeader(2,1) = "Shipping Code"
	aSortHeader(3,1) = "Shipping Rate"
	aSortHeader(4,1) = "Active"

	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 1 to 4
		If cInt(pstrOrderBy) = i Then
			If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
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
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & mstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
		End If
	next 'i

	Response.Write "<tr><td colspan=4>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary' id='tblSummary' " _
				 & ">"
    .Write "<COLGROUP align='left' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"
    .Write "<COLGROUP align='center' width='25%'>"

	if pRS.RecordCount > 0 then
		pRS.MoveFirst
		for i=1 to prs.RecordCount
			pstrTitle = "Click to edit " & prs("shipMethod") & "."
			pstrURL = "sfShippingAdmin.asp?Action=View&ID=" & pRS("shipID")

			if (trim(pRS("shipID")) = plngID) then
        		Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
			else
				if cBool(pRS("shipIsActive")) then
	                Response.Write "<TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
				else
	                Response.Write "<TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
        		end if
        	end if

        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & pRS("shipMethod") & "</a></TD>"
        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & pRS("shipCode") & "</TD>"
        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & pRS("shipRates") & "</TD>"
			if ConvertToBoolean(pRS("shipIsActive"), False) then
        		Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='sfShippingAdmin.asp?Action=Deactivate&ID=" & pRS("shipID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prs("shipMethod") & "'>Active</a></TD></TR>" & vbCrLf
			else
        		Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='sfShippingAdmin.asp?Action=Activate&ID=" & pRS("shipID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prs("shipMethod") & "'>Inactive</a></TD></TR>" & vbCrLf
        	end if
			prs.MoveNext
		next
	else
		.Write "<TR><TD colspan=3><h3>There are no shipping methods</h3></TD></TR>"	
	end if
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	
End Sub

End Class   'clsPromotionshipping

Sub LoadFilter

dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")

'Build Filter

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Shipping Method
			pstrOrderBy = "shipMethod"
		Case "2"	'Code
			pstrOrderBy = "shipCode"
		Case "3"	'Rate
			pstrOrderBy = "shipRates"
		Case "4"	'Active
			pstrOrderBy = "shipIsActive"
	End Select	

	If len(pstrOrderBy) > 0 then
		mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & mstrSortOrder
	Else
		mstrsqlWhere = ""
	End If
	
End Sub    'LoadFilter

mstrPageTitle = "Supported Shipping Methods Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsShipping
Dim mstrSortOrder, mstrOrderBy, mstrsqlWhere
Dim mlngID

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    mlngID = Request.QueryString("ID")
    if len(mlngID) = 0 Then mlngID = Request.Form("ID")
    
    Call LoadFilter
    Set mclsShipping = New clsShipping

With mclsShipping    
    Select Case mAction
        Case "New", "Update"
            If .Update then
				If .LoadAll Then .Find mlngID
			Else
				.LoadAll
				.LoadFromRequest
			End If
        Case "Delete"
            .Delete mlngID
            .LoadAll
        Case "View"
            If .LoadAll Then .Find mlngID
         Case "Activate", "Deactivate"
			.Activate mlngID, mAction="Activate"
            If .LoadAll Then .Find mlngID
       Case Else
            .LoadAll
    End Select

Call WriteHeader("body_onload();",True)
%>
<HTML>
<HEAD>
<TITLE>shipping Administration</TITLE>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var strDetailTitle = "Edit <%= .METHODS %>";

function body_onload()
{
	theDataForm = document.frmData;
}

function SetDefaults()
{
    theDataForm.ACTIVE.checked = false;
    theDataForm.EDIT.checked = true;
    theDataForm.CODE.value = "";
    theDataForm.ID.value = "";
    theDataForm.METHODS.value = "";
    theDataForm.RATES.value = "";
return(true);
}

function btnNew_onclick()
{
    SetDefaults();
    theDataForm.btnUpdate.value = "Add Shipping Method";
    theDataForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
    theDataForm.METHODS.focus();
}

function btnDelete_onclick()
{
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + frmData.METHODS.value + "?");
    if (blnConfirm)
    {
    theDataForm.Action.value = "Delete";
    theDataForm.submit();
    return(true);
    }
    {
    return(false);
    }
}

function btnReset_onclick()
{
    theDataForm.btnUpdate.value = "Save Changes";
    theDataForm.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function ValidInput(theForm)
{

  if (isEmpty(theForm.METHODS,"Please enter a name for the shipping method.")) {return(false);}
  if (isEmpty(theForm.CODE,"Please enter a code for the shipping method.")) {return(false);}
  if (!isNumeric(theForm.RATES,false,"Please enter a value for the Rate.")) {return(false);}

    return(true);
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
</HEAD>
<BODY onload="body_onload();">
<CENTER>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<% .OutputSummary %>

<FORM action='sfShippingAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=ID name=ID value=<%= .ID %>>
<input type=hidden id=EDIT name=EDIT value=<%= .EDIT %>>
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=OrderBy name=OrderBy value='<%= mstrOrderBy %>'>
<input type=hidden id=SortOrder name=SortOrder value='<%= mstrSortOrder %>'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit <%= .METHODS %></span></th>
  </tr>

      <tr>
        <TD class="label"><LABEL id=lblMETHODS for=METHODS>Method:</LABEL></TD>
        <TD><INPUT id=METHODS name=METHODS Value="<%= .METHODS %>" maxlength=50 size=50></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lblCODE for=CODE>Code:</LABEL></TD>
        <TD><INPUT id=CODE name=CODE Value="<%= .CODE %>" maxlength=25 size=25></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lblRATES for=RATES>Rate:</LABEL></TD>
        <TD><INPUT id=RATES name=RATES Value="<%= .RATES %>" maxlength=10 size=10></TD>
      </tr>

      <TR>
        <TD>&nbsp;</TD>
        <TD><INPUT id=ACTIVE name=ACTIVE type=checkbox <% if .ACTIVE then Response.Write "checked" %>><LABEL id=lblACTIVE for=ACTIVE>Check if Active</LABEL>
      </TR>
  <TR>
    <TD colspan=2 align=center>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick()'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=Reset value=Reset onclick='return btnReset_onclick()'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick()'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</FORM>

</CENTER>
</BODY>
</HTML>
<%
End With

Set mclsShipping = Nothing
Set cnn = Nothing

Response.Flush
%>
