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

Class clsColor
	'Assumptions:
	'   cnn: defines a previously opened connection to the database

	'class variables
	Private cstrDelimeter
	Private pstrMessage
	Private pRS
	Private pblnError
	'database variables

	Private plngID
	Private pstrCOLOR
	Private pstrCOLOR_CODE

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


	Public Property Get ID()
	    ID = plngID
	End Property

	Public Property Get COLOR()
	    COLOR = pstrCOLOR
	End Property

	Public Property Get COLOR_CODE()
	    COLOR_CODE = pstrCOLOR_CODE
	End Property

	'***********************************************************************************************

	Public Function Find(lngID)

	'On Error Resume Next

	    with pRS
			if .RecordCount > 0 then
				.MoveFirst
				If len(lngID) <> 0 then
					.Find "slctvalID=" & lngID
				Else
					.Find "slctvalID=" & plngID
				End If
				if not .EOF then LoadValues(pRS)
			end if
		end with
	  
	End Function    'Load

	'***********************************************************************************************

	Private Sub LoadValues(rs)

	    plngID = trim(rs("slctvalID"))
	    pstrCOLOR = trim(rs("slctvalColor"))
	    pstrCOLOR_CODE = trim(rs("slctvalColorCode"))

	End Sub 'LoadValues

	Public Sub LoadFromRequest

	    With Request.Form
	        plngID = Trim(.Item("ID"))
	        pstrCOLOR = Trim(.Item("COLOR"))
	        pstrCOLOR_CODE = Trim(.Item("COLOR_CODE"))
	    End With

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Public Function Load(lngID)

	Dim sql
	Dim rs

	'On Error Resume Next

	    sql = "Select slctvalID,slctvalColor,slctvalColorCode from sfSelectValues where slctvalID = " & lngID
	    Set rs = server.CreateObject("adodb.Recordset")
	    Set rs = cnn.Execute(sql)
	    If Not (rs.EOF Or rs.BOF) Then
	        Call LoadValues(rs)
	        Load = True
	    End If

	    rs.Close
	    Set rs = Nothing

	End Function    'Load

	'***********************************************************************************************

	Public Function LoadAll()

	On Error Resume Next

	    Set pRS = server.CreateObject("adodb.Recordset")

	    With pRS

	        .ActiveConnection = cnn
	        .CursorLocation = 2 'adUseClient
	        .CursorType = 3 'adOpenStatic
	        .LockType = 1 'adLockReadOnly
	        .Source = "Select slctvalID,slctvalColor,slctvalColorCode from sfSelectValues where slctvalColor<>'' Order By slctvalColor"
	        .open

	        If Not (.EOF Or .BOF) Then
	            Call LoadValues(pRS)
	            LoadAll = True
	        End If

	    End With

	End Function    'LoadAll

	'***********************************************************************************************

	Public Function Delete(lngID)

	Dim sql
	Dim rs

	On Error Resume Next


	    sql = "Update sfSelectValues Set slctvalColor=Null, slctvalColorCode=Null where slctvalID = " & lngID
	    cnn.Execute sql, , 128
	    If (Err.Number = 0) Then
	        pstrMessage = "The color was successfully deleted."
	        Delete = True
	    Else
	        pstrMessage = Err.Description
	        Delete = False
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
			blnAdd = False
	        If Len(plngID) = 0 Then plngID = 0

	        sql = "Select slctvalID,slctvalColor,slctvalColorCode from sfSelectValues"
	        Set rs = server.CreateObject("adodb.Recordset")
	        rs.open sql, cnn, 1, 3
	        If rs.EOF Then
	            rs.AddNew
	            blnAdd = True
	        Else
				If plngID = 0 then
				    blnAdd = True
					rs.Find "slctvalColor=Null"
				Else
					rs.Find "slctvalID=" & plngID
				End If
				If rs.EOF then
				    blnAdd = True
					rs.AddNew
				End If
	        End If

	        rs("slctvalColor") = pstrCOLOR
	        rs("slctvalColorCode") = pstrCOLOR_CODE

	        rs.Update

	        If Err.Number = -2147217887 Then
	            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
	                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
	                pblnError = True
	            End If
	        ElseIf Err.Number <> 0 Then
	            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
	        End If
	        
	        plngID = rs("slctvalID")

	        rs.Close
	        Set rs = Nothing
	        
	        If Err.Number = 0 Then
	            If blnAdd Then
	                pstrMessage = "The color " & pstrCOLOR & " was successfully added."
	            Else
	                pstrMessage = "The color " & pstrCOLOR & " was successfully updated."
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

	    If Len(pstrCOLOR) = 0 Then
	        strError = strError & "Please enter a color." & cstrDelimeter
	    End If

	    If Len(pstrCOLOR_CODE) = 0 Then
	        strError = strError & "Please enter a color code." & cstrDelimeter
	    End If

	    pstrMessage = strError
	    ValidateValues = (Len(strError) = 0)

	End Function 'ValidateValues

	Sub OutputSummary

	dim i

	With Response
    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='50%'>"
    .Write "<COLGROUP align='left' width='50%'>"
    .Write "  <tr class='tblhdr'>"
	.Write "  <TH>&nbsp;&nbsp;&nbsp;Color</TH>" & vbCrLf
	.Write "  <TH>Color Code</TH>" & vbCrLf
	.Write "  </TR>" & vbCrLf
	Response.Write "<tr><td colspan=2>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
		& "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary' id='tblSummary' " _
		& ">"
    Response.Write "<COLGROUP align='left' width='50%'>"
    Response.Write "<COLGROUP align='left' width='50%'>"
		if pRS.RecordCount > 0 then
			pRS.MoveFirst
			for i=1 to prs.RecordCount
				if Trim(pRS("slctvalID")) = plngID then
	                Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);'>"
				else
	                Response.Write "<TR 'title='Click to edit " & pRS("slctvalColor") & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('sfColorAdmin.asp?Action=View&ID=" & pRS("slctvalID") & "')" & chr(34) & ">"
	        	end if
	        	
	        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;<a href='sfColorAdmin.asp?Action=View&ID=" & pRS("slctvalID") & "' title='Click to edit'>" & pRS("slctvalColor") & "</a>&nbsp;&nbsp;</TD>"
	        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & pRS("slctvalColorCode") & "</TD></TR>"
				prs.MoveNext
			next
		else
			Response.Write "<TR><TD colspan=2><h3>There are no Colors</h3></TD></TR>"	
		end if
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	End Sub

End Class   'clsColor

mstrPageTitle = "Color Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsColor
Dim mlngID

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclsColor = New clsColor
With mclsColor    
    Select Case mAction
        Case "New", "Update"
            If .Update then
				mlngID = .ID
				If .LoadAll Then .Find mlngID
			Else
				.LoadAll
				.LoadFromRequest
			End If
        Case "Delete"
            .Delete Request.Form("ID")
            .LoadAll
        Case "View"
            If .LoadAll Then .Find Request.QueryString("ID")
       Case Else
            .LoadAll
    End Select

	Call WriteHeader("body_onload();",True)
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var strDetailTitle = "Edit <%= .COLOR %>";

function body_onload()
{
	theDataForm = document.frmData;
}

function SetDefaults()
{
    theDataForm.ID.value = "";
    theDataForm.COLOR.value = "";
    theDataForm.COLOR_CODE.value = "";
return(true);
}

function btnNew_onclick()
{
    SetDefaults();
    theDataForm.btnUpdate.value = "Add Color";
    theDataForm.btnDelete.disabled = true;
    theDataForm.COLOR.focus();
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick()
{
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + frmData.COLOR.value + "?");
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
    theDataForm.Action.value = "Update";
    theDataForm.btnUpdate.value = "Save Changes";
    frmData.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function ValidInput(theForm)
{
  if (isEmpty(theForm.COLOR,"Please enter a color.")) {return(false);}
  if (isEmpty(theForm.COLOR_CODE,"Please enter a color code.")) {return(false);}

  return(true);
}

//-->
</SCRIPT>
<BODY onload="body_onload();">
<center>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<% .OutputSummary %>
<FORM action='sfColorAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=ID name=ID value=<%= .ID %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit <%= .COLOR %></span></th>
  </tr>
      <tr>
        <TD class="label"><LABEL id=lblCOLOR for=COLOR>Color:</LABEL></TD>
        <TD>&nbsp;<INPUT id=COLOR name=COLOR Value="<%= .COLOR %>" maxlength=50 size=50></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lblCOLOR_CODE for=COLOR_CODE>Color Code:</LABEL></TD>
        <TD>&nbsp;<INPUT id=COLOR_CODE name=COLOR_CODE Value="<%= .COLOR_CODE %>" maxlength=50 size=50></TD>
      </tr>
  <TR>
    <TD colspan=2 align=center>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick()'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=Reset value=Reset onclick='return btnReset_onclick()'>&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick()'>&nbsp;
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</FORM>
</center>
</BODY>
</HTML>
<%
End With

Set mclsColor = Nothing
Set cnn = Nothing

Response.Flush
%>
