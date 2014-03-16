<% Option Explicit 
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version   1.0                                                      *
'*   Release Date      April 13, 2001			                                *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsPricingLevel
	'Assumptions:
	'   cnn: defines a previously opened connection to the database

	'class variables
	Private cstrDelimeter
	Private pstrMessage
	Private pRS
	Private pblnError
	'database variables

	Private plngID
	Private pstrPricingLevelName
	Private pstrPricingLevelNotes

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

	Public Property Get PricingLevelName()
	    PricingLevelName = pstrPricingLevelName
	End Property

	Public Property Get PricingLevelNotes()
	    PricingLevelNotes = pstrPricingLevelNotes
	End Property

	'***********************************************************************************************

	Public Function Find(lngID)

	'On Error Resume Next

	    with pRS
			if .RecordCount > 0 then
				.MoveFirst
				If len(lngID) <> 0 then
					.Find "PricingLevelID=" & lngID
				Else
					.Find "PricingLevelID=" & plngID
				End If
				if not .EOF then LoadValues(pRS)
			end if
		end with
	  
	End Function    'Load

	'***********************************************************************************************

	Private Sub LoadValues(rs)

	    plngID = trim(rs("PricingLevelID"))
	    pstrPricingLevelName = trim(rs("PricingLevelName"))
	    pstrPricingLevelNotes = trim(rs("PricingLevelNotes"))

	End Sub 'LoadValues

	Public Sub LoadFromRequest

	    With Request.Form
	        plngID = Trim(.Item("ID"))
	        pstrPricingLevelName = Trim(.Item("PricingLevelName"))
	        pstrPricingLevelNotes = Trim(.Item("PricingLevelNotes"))
	    End With

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Public Function Load(lngID)

	Dim sql
	Dim rs

	'On Error Resume Next

	    sql = "Select PricingLevelID,PricingLevelName,PricingLevelNotes from PricingLevels where PricingLevelID = " & lngID
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
	        .Source = "Select PricingLevelID,PricingLevelName,PricingLevelNotes from PricingLevels where PricingLevelName<>'' Order By PricingLevelName"
	        .open

			If Err.number <> 0 Then
				Response.Write "<div class='FatalError'>You need to upgrade your database to use Pricing Level Manager</div>" _
							   & "<h3><a href='ssInstallationPrograms/PricingLevelSF5Upgrade.asp'>Click here to upgrade</a></h3>"
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				Response.Flush
				LoadAll = False
				Exit Function
			End If
			
	        If Not (.EOF Or .BOF) Then
	            Call LoadValues(pRS)
	            LoadAll = True
	        End If

	    End With

	End Function    'LoadAll

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

	        sql = "Select PricingLevelID,PricingLevelName,PricingLevelNotes from PricingLevels"
	        Set rs = server.CreateObject("adodb.Recordset")
	        rs.open sql, cnn, 1, 3
	        If rs.EOF Then
	            rs.AddNew
	            blnAdd = True
	        Else
				If plngID = 0 then
				    blnAdd = True
					rs.Find "PricingLevelName=Null"
				Else
					rs.Find "PricingLevelID=" & plngID
				End If
				If rs.EOF then
				    blnAdd = True
					rs.AddNew
				End If
	        End If

	        rs("PricingLevelName") = pstrPricingLevelName
	        rs("PricingLevelNotes") = pstrPricingLevelNotes

	        rs.Update

	        If Err.Number = -2147217887 Then
	            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
	                pstrMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
	                pblnError = True
	            End If
	        ElseIf Err.Number <> 0 Then
	            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
	        End If
	        
	        plngID = rs("PricingLevelID")

	        rs.Close
	        Set rs = Nothing
	        
	        If Err.Number = 0 Then
	            If blnAdd Then
	                pstrMessage = "The pricing level " & pstrPricingLevelName & " was successfully added."
	            Else
	                pstrMessage = "The pricing level " & pstrPricingLevelName & " was successfully updated."
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

	    If Len(pstrPricingLevelName) = 0 Then
	        strError = strError & "Please enter a pricing level." & cstrDelimeter
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
	.Write "  <TH>&nbsp;&nbsp;&nbsp;Pricing Level</TH>" & vbCrLf
	.Write "  <TH>Notes</TH>" & vbCrLf
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
				if Trim(pRS("PricingLevelID")) = plngID then
	                Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);'>"
				else
	                Response.Write "<TR 'title='Click to edit " & pRS("PricingLevelName") & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('ssPricingLevelAdmin.asp?Action=View&ID=" & pRS("PricingLevelID") & "')" & chr(34) & ">"
	        	end if
	        	
	        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;<a href='ssPricingLevelAdmin.asp?Action=View&ID=" & pRS("PricingLevelID") & "' title='Click to edit'>" & pRS("PricingLevelName") & "</a>&nbsp;&nbsp;</TD>"
	        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & pRS("PricingLevelNotes") & "</TD></TR>"
				prs.MoveNext
			next
		else
			Response.Write "<TR><TD colspan=2><h3>There are no pricing levels</h3></TD></TR>"	
		end if
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	End Sub

End Class   'clsPricingLevel

mstrPageTitle = "Pricing Level Administration"

%>
<!--#include file="SSLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="SSLibrary/modDatabase.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsPricingLevel
Dim mlngID

	Call WriteHeader("body_onload();",True)

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclsPricingLevel = New clsPricingLevel
	With mclsPricingLevel    
		Select Case mAction
			Case "New", "Update"
				If .Update then
					mlngID = .ID
					If .LoadAll Then .Find mlngID
				Else
					.LoadAll
					.LoadFromRequest
				End If
			Case "View"
				If .LoadAll Then .Find Request.QueryString("ID")
		Case Else
				.LoadAll
		End Select
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var strDetailTitle = "Edit <%= .PricingLevelName %>";

function body_onload()
{
	theDataForm = document.frmData;
}

function SetDefaults()
{
    theDataForm.ID.value = "";
    theDataForm.PricingLevelName.value = "";
    theDataForm.PricingLevelNotes.value = "";
return(true);
}

function btnNew_onclick()
{
    SetDefaults();
    theDataForm.btnUpdate.value = "Add Pricing Level";
    theDataForm.PricingLevelName.focus();
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnReset_onclick()
{
    theDataForm.Action.value = "Update";
    theDataForm.btnUpdate.value = "Save Changes";
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function ValidInput(theForm)
{
  if (isEmpty(theForm.PricingLevelName,"Please enter a pricing level.")) {return(false);}

  return(true);
}

//-->
</SCRIPT>
<BODY onload="body_onload();">
<center>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<% .OutputSummary %>
<FORM action='ssPricingLevelAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=ID name=ID value=<%= .ID %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit <%= .PricingLevelName %></span></th>
  </tr>
      <tr>
        <TD class="label"><LABEL id=lblPricingLevelName for=PricingLevelName>Pricing Level:</LABEL></TD>
        <TD>&nbsp;<INPUT id=PricingLevelName name=PricingLevelName Value="<%= .PricingLevelName %>" maxlength=20 size=20></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lblPricingLevelNotes for=PricingLevelNotes>Notes:</LABEL></TD>
        <TD>&nbsp;<INPUT id=PricingLevelNotes name=PricingLevelNotes Value="<%= .PricingLevelNotes %>" maxlength=255 size=50></TD>
      </tr>
  <TR>
    <TD colspan=2 align=center>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick()'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=Reset value=Reset onclick='return btnReset_onclick()'>&nbsp;
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

Set mclsPricingLevel = Nothing
Set cnn = Nothing

Response.Flush
%>
