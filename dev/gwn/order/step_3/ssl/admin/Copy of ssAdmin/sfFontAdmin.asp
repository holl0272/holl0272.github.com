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

Class clsFont
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private plngID
Private pstrslctvalFontType

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

Public Property Get slctvalFontType()
    slctvalFontType = pstrslctvalFontType
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
				.MoveLast
			End If
			if not .EOF then LoadValues(pRS)
		end if
	end with
  
End Function    'Load

'***********************************************************************************************

Private Sub LoadValues(rs)

    plngID = rs("slctvalID")
    pstrslctvalFontType = trim(rs("slctvalFontType"))

End Sub 'LoadValues

Public Sub LoadFromRequest

    With Request.Form
        plngID = Trim(.Item("ID"))
        pstrslctvalFontType = Trim(.Item("slctvalFontType"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Load(lngID)

Dim sql
Dim rs

'On Error Resume Next

    sql = "Select slctvalID,slctvalFontType from sfSelectValues where slctvalID = " & lngID
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

'On Error Resume Next

    Set pRS = server.CreateObject("adodb.Recordset")

    With pRS

        .ActiveConnection = cnn
        .CursorLocation = 2 'adUseClient
        .CursorType = 3 'adOpenStatic
        .LockType = 1 'adLockReadOnly
        .Source = "Select slctvalID,slctvalFontType from sfSelectValues where slctvalFontType<>'' Order By slctvalFontType"
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

	sql = "Update sfSelectValues Set slctvalFontType=Null where slctvalID = " & lngID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "The font was successfully deleted."
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

        sql = "Select slctvalID,slctvalFontType from sfSelectValues"
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
	    If rs.EOF Then
	        rs.AddNew
            blnAdd = True
        Else
			If plngID = 0 then
			    blnAdd = True
				rs.Find "slctvalFontType=Null"
			Else
				rs.Find "slctvalID = " & plngID
			End If
			If rs.EOF then
			    blnAdd = True
				rs.AddNew
			End If
        End If

        rs("slctvalFontType") = pstrslctvalFontType
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
                pstrMessage = "The Font " & pstrslctvalFontType & " was successfully added."
            Else
                pstrMessage = "The Font " & pstrslctvalFontType & " was successfully updated."
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

    If Len(pstrslctvalFontType) = 0 Then
        strError = strError & "Please enter a Font." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)

End Function 'ValidateValues

Sub OutputSummary

dim i
dim j
Dim pstrTitle, pstrURL
Dim pstrCell

	Response.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
		 & "bgcolor='whitesmoke' name='tblSummary' id='tblSummary' " _
		 & ">" & vbcrlf
	Response.Write "<COLGROUP align='left'>" & vbcrlf
	Response.Write "<COLGROUP align='left'>" & vbcrlf
	Response.Write "<COLGROUP align='left'>" & vbcrlf
	Response.Write "<COLGROUP align='left'>" & vbcrlf
	if pRS.RecordCount > 0 then
		pRS.MoveFirst
		j=1

		for i=1 to prs.RecordCount
			pstrTitle = "Click to edit " & pRS("slctvalFontType")
			pstrURL = "sfFontAdmin.asp?Action=View&ID=" & pRS("slctvalID")
			if pRS("slctvalID") = plngID then
				pstrCell = "  <TD class='Selected'>&nbsp;&nbsp;<a href='" & pstrURL & "' title='" & pstrTitle & "' onmouseover='return window.status=this.title;' onmouseout=""window.status='';"">" & pRS("slctvalFontType") & "</a>&nbsp;&nbsp;</TD>" & vbcrlf
			Else
				pstrCell = "  <TD>&nbsp;&nbsp;<a href='" & pstrURL & "' title='" & pstrTitle & "' onmouseover='return window.status=this.title;' onmouseout=""window.status='';"">" & pRS("slctvalFontType") & "</a>&nbsp;&nbsp;</TD>" & vbcrlf
			End If
			if j=4 then
        		Response.Write pstrCell
        		Response.Write "</TR>" & vbcrlf
				j=1
			elseif j=1 then
       			Response.Write "<TR>" & vbcrlf
       			Response.Write pstrCell
				j=j+1
			else
       			Response.Write pstrCell
				j=j+1
			end if
			prs.MoveNext
		next
		if j=2 then
        	Response.Write "<TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>" & vbcrlf
		elseif j=3 then
        	Response.Write "<TD>&nbsp;</TD><TD>&nbsp;</TD></TR>" & vbcrlf
		elseif j=4 then
        	Response.Write "<TD>&nbsp;</TD></TR>" & vbcrlf
		end if
	else
		Response.Write "<TR><TD colspan=4><h3>There are no Fonts</h3></TD></TR>" & vbcrlf
	end if
    Response.Write "</TABLE>" & vbcrlf

End Sub

End Class   'clsFont

mstrPageTitle = "Font Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsFont
Dim mlngID

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclsFont = New clsFont
With mclsFont    
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
<HTML>
<HEAD>
<TITLE>Font Administration</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var strDetailTitle = "Edit <%= .slctvalFontType %>";

function body_onload()
{
	theDataForm = document.frmData;
}

function SetDefaults()
{
    theDataForm.ID.value = "";
    theDataForm.slctvalFontType.value = "";
return(true);
}

function btnNew_onclick()
{
    SetDefaults();
    theDataForm.btnUpdate.value = "Add Font";
    theDataForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick()
{
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + theDataForm.slctvalFontType.value + "?");
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

function ValidInput()
{
  if (theDataForm.slctvalFontType.value == "")
  {
    alert("Please enter a Font.")
    theDataForm.slctvalFontType.focus();
    return(false);
  }
    return(true);
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="body_onload();">
<CENTER>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<% .OutputSummary %>

<FORM action='sfFontAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=ID name=ID value=<%= .ID %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="center">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit <%= .slctvalFontType %></span></th>
  </tr>
      <tr>
        <TD><LABEL id=lblslctvalFontType for=slctvalFontType>Font:</LABEL>&nbsp;<INPUT id=slctvalFontType name=slctvalFontType Value="<%= .slctvalFontType %>" maxlength=50 size=50></TD>
      </tr>
  <TR>
    <TD align='center'>
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

Set mclsFont = Nothing
Set cnn = Nothing

Response.Flush
%>
