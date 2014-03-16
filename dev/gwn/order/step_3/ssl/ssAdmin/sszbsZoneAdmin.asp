<%Option Explicit
'********************************************************************************
'*   Webstore Manager Gold                                                      *
'*   Release Version   1.0                                                      *
'*   Release Date      January 15, 2001                                         *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsZone
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsZone
Private pblnError

'database variables
Private plngZoneID
Private pstrZoneName
Private pstrZoneCountries
Private pstrZoneStates
Private pstrZoneZIPs

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
    Set prsZone = Nothing
End Sub

'***********************************************************************************************

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

'***********************************************************************************************

Public Property Get ZoneID()
    ZoneID = plngZoneID
End Property
Public Property Get ZoneName()
    ZoneName = pstrZoneName
End Property
Public Property Get ZoneCountries()
    ZoneCountries = pstrZoneCountries
End Property
Public Property Get ZoneStates()
    ZoneStates = pstrZoneStates
End Property
Public Property Get ZoneZIPs()
    ZoneZIPs = pstrZoneZIPs
End Property

'***********************************************************************************************

Private Sub ClearValues()

	plngZoneID = ""
	pstrZoneName = ""
	pstrZoneCountries = ""
	pstrZoneStates = ""
	pstrZoneZIPs = ""

End Sub 'ClearValues

Private Sub LoadValues

	plngZoneID = Trim(prsZone("ZoneID"))
	pstrZoneName = Trim(prsZone("ZoneName"))
	
	If not isNull(prsZone("ZoneCountries")) Then 
		pstrZoneCountries = StripDelimeters(Trim(prsZone("ZoneCountries")))
	Else
		pstrZoneCountries = ""
	End If
	If not isNull(prsZone("ZoneStates")) Then
		pstrZoneStates = StripDelimeters(Trim(prsZone("ZoneStates")))
	Else
		pstrZoneStates = ""
	End If
	If not isNull(prsZone("ZoneZIPs")) Then
		pstrZoneZIPs = StripDelimeters(Trim(prsZone("ZoneZIPs")))
	Else
		pstrZoneZIPs = ""
	End If
	
End Sub 'LoadValues

Private Function StripDelimeters(strSource)

	if left(strSource,1) = ";" Then strSource = right(strSource,len(strSource)-1)
	if right(strSource,1) = ";" Then strSource = left(strSource,len(strSource)-1)
	StripDelimeters = strSource

End Function

Private Sub LoadFromRequest

    With Request.Form
		plngZoneID = Trim(.Item("ZoneID"))
		pstrZoneName = Trim(.Item("ZoneName"))
		pstrZoneCountries = Trim(.Item("ZoneCountry"))
		pstrZoneStates = Trim(.Item("ZoneState"))
		pstrZoneZIPs = Trim(.Item("ZoneZIP"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function FindZone(lngZoneID)

'On Error Resume Next

    With prsZone
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngZoneID) <> 0 Then
                .Find "ZoneID=" & lngZoneID
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues
        End If
    End With

End Function    'FindZone

'***********************************************************************************************

Public Function Load()

'On Error Resume Next

    Set prsZone = GetRS("Select * from ssShipZones Order By ZoneName")

	If prsZone.State <> 1 Then
		Response.Write "<h3><font color=red>The Zone Based Shipping add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
		Response.Write "<a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=ZBS'>Click here to upgrade</a></h3>"
		Response.Flush
		Err.Clear
		Load = False
		Exit Function
	End If

    If not prsZone.EOF Then LoadValues
    Load = (Not prsZone.EOF)

End Function    'Load

'***********************************************************************************************

Public Function DeleteZone(lngID)

Dim sql

'On Error Resume Next

	If len(lngID) = 0 Then Exit Function
    sql = "Delete from ssShipZones where ZoneID = " & lngID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "The Zone was successfully deleted."
        DeleteZone = True
    Else
        pstrMessage = Err.Description
        DeleteZone = False
    End If

End Function    'DeleteZone

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
        If Len(plngZoneID) = 0 Then plngZoneID = 0

        sql = "Select * from ssShipZones where ZoneID = " & plngZoneID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("ZoneName") = pstrZoneName
		if len(pstrZoneCountries) > 0 Then rs("ZoneCountries") = ";" & pstrZoneCountries & ";"
		if len(pstrZoneStates) > 0 Then rs("ZoneStates") = ";" & pstrZoneStates & ";"
		if len(pstrZoneZIPs) > 0 Then rs("ZoneZIPs") = ";" & pstrZoneZIPs & ";"
		
		rs.Update
        plngZoneID = rs("ZoneID")
        rs.Close

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create dupliZonee values in the index, primary key, or relationship.  Change the data in the field or fields that contain dupliZonee data, remove the index, or redefine the index to permit dupliZonee entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number = -2147217873 Then
            If Err.Description = "[Microsoft][ODBC Microsoft Access Driver]Error in row " Then
                pstrMessage = "<H4>You have already created a shipping zone with this name.<br />Please enter a different shipping zone name.</H4><br />"
            Else
				pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
            End If
        ElseIf Err.Number = 3219 Then
            If Err.Description = "Operation is not allowed in this context." Then
                pstrMessage = "<H4>You have already created a shipping zone with this name.<br />Please enter a different shipping zone name.</H4><br />"
            Else
				pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & ": - :" & Err.Description & ":<br />"
        End If
        
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrZoneName & " was successfully added."
            Else
                pstrMessage = "The changes to " & pstrZoneName & " were successfully saved."
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

Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim pstrTitle, pstrURL, pstrAbbr
Dim pstrSelect
Dim pblnActive

	With Response

		.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke' id='tblSummary'>"
		.Write "<colgroup align='left'>" & vbCrLf
		.Write "<tr class='tblhdr'>" & vbCrLf
	    .Write "  <TH align=left width='30%'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Zone ID</TH>" & vbCrLf
	    .Write "  <TH align=left width='70%'>Zone</TH>" & vbCrLf
		.Write "</tr>" & vbCrLf

		.Write "<tr><td colspan=2>" & vbCrLf
		.Write "<div name='divSummary' style='height:100; overflow:auto;'>"
		.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
				 & ">"
		.Write "<colgroup align='left' width='5%'>" & vbcrlf
		.Write "<colgroup align='left' width='25%'>" & vbcrlf
		.Write "<colgroup align='left' width='70%'>" & vbcrlf
    If prsZone.RecordCount > 0 Then
        prsZone.MoveFirst
        For i = 1 To prsZone.RecordCount
			pstrURL = "sszbsZoneAdmin.asp?Action=View&ZoneID=" & prsZone("ZoneID")
			pstrTitle = "Click to view " & prsZone("ZoneName")
			pstrSelect = "title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();'"
			
            If Trim(prsZone("ZoneID")) = plngZoneID Then
                .Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>" & vbcrlf
        		.Write "<TD>&nbsp;</TD>" & vbcrlf
        		.Write "<TD>" & prsZone("ZoneID") & "&nbsp;</TD>" & vbcrlf
            Else
				.Write "<TR " & pstrSelect & " onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
        		.Write "<TD>&nbsp;</TD>" & vbcrlf
        		.Write "<TD><a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prsZone("ZoneID") & "&nbsp;</a></TD>" & vbcrlf
            End If
            
       		.Write "<TD >" & prsZone("ZoneName") & "</TD>" & vbcrlf
			.Write "</TR>" & vbcrlf
            prsZone.MoveNext
        Next
    Else
			.Write "<TR><TD align=center><h3>There are no Zones</h3></TD></TR>"
    End If
		.Write "</td></tr></TABLE></div>"
		.Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrZoneName) = 0 Then
        strError = strError & "Please enter a Zone name." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues

'***********************************************************************************************

End Class   'clsZone

'--------------------------------------------------------------------------------------------------
%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'**************************************************
'
'	Start Code Execution
'

Dim mclsZone
Dim mAction
Dim mlngZoneID, mlngsubZoneID

	'call InitializeConnection(cnn)
	mstrPageTitle = "Zone Administration"

    mlngZoneID = Request.QueryString("ZoneID")
    If len(mlngZoneID) = 0 Then mlngZoneID = Request.Form("ZoneID")

    mlngsubZoneID = Request.QueryString("subZoneID")
    If len(mlngsubZoneID) = 0 Then mlngsubZoneID = Request.Form("subZoneID")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclsZone = New clsZone
    With mclsZone
    
    Select Case mAction
        Case "New", "Update"
            .Update
            If .Load Then .FindZone mlngZoneID
        Case "DeleteZone"
			.DeleteZone mlngZoneID
            .Load
        Case "View"
			If len(mlngZoneID) > 0 Then
				If .Load Then .FindZone mlngZoneID
			Else
				If .Load Then .FindsubZone mlngsubZoneID
			End If
        Case Else
            .Load
    End Select
    
Call WriteHeader("body_onload();",True)
%>

<SCRIPT LANGUAGE=javascript>
<!--
var strDetailTitle = "<%= .ZoneName %> Details";
var c_strDelimeter = ";";

var theForm;
var strZoneID;
var strZIP;
var strCountry;
var strState;

var pblnAddSub;

function body_onload()
{
	theForm = frmData;
	strZoneID = theForm.ZoneID.value;
	strZIP = theForm.ZoneZip.value;
	strCountry = theForm.ZoneCountry.value;
	strState = theForm.ZoneState.value;

	InitializeLists();
	pblnAddSub = false;
}

function btnNewZone_onclick()
{

	pblnAddSub = false;

	MoveAll(theForm.targetZoneCountry,theForm.sourceZoneCountry)
	MoveAll(theForm.targetZoneState,theForm.sourceZoneState)

    theForm.ZoneID.value = "";
    theForm.ZoneName.value = "";

    theForm.btnUpdate.value = "Add Zone";
    theForm.btnDeleteZone.disabled = true;
	theForm.btnUpdate.disabled = false;
	theForm.btnReset.disabled = false;

    theForm.ZoneName.focus();
    document.all("spanZoneName").innerHTML = theForm.btnUpdate.value;

}

function InitializeLists()
{

var aryCountries = strCountry.split(c_strDelimeter);
var aryStates = strState.split(c_strDelimeter);

for (var i=0; i<aryCountries.length; i++)
{
	isInList(aryCountries[i],theForm.sourceZoneCountry,true);
}
for (var i=0; i<aryStates.length; i++)
{
	isInList(aryStates[i],theForm.sourceZoneState,true);
}
	MoveItems("Country",true)
	MoveItems("State",true)

}

function isInList(theValue,theSelect,blnHighlight)
{

	if (theSelect.length == 0){return(false);}

	for (var i=0; i < theSelect.length;i++)
	{
		if (theSelect.options[i].value == theValue)
		{
			if (blnHighlight) {theSelect.options[i].selected = true}
			return(true);
		}
	}
	return(false);
}

function MoveItems(theOption,blnAdd)
{
var theSource;
var theTarget;
var intSelected;

	theSource = eval("theForm.sourceZone" + theOption);
	theTarget = eval("theForm.targetZone" + theOption);
	
	if (blnAdd)
	{
		if (theSource.selectedIndex == -1){return false;};
		MoveItemsToo(theSource,theTarget);
	}else{
		if (theTarget.selectedIndex == -1){return false;};

		intSelected = theTarget.selectedIndex;
		MoveItemsToo(theTarget,theSource);
	}
	
}

function MoveItemsToo(theSource,theTarget)
{
var j = theTarget.length;

	for (var i=0; i < theSource.length;i++)
	{
		if (theSource.options[i].selected)
		{
			theTarget.options[j] = new Option(theSource.options[i].text, theSource.options[i].value);
			j++;
		}
	}
	for (var i=0; i < theSource.length;i++)
	{
		if (theSource.options[i].selected){theSource.options.remove(i); i--;}
	}

}

function MoveAll(theSource,theTarget)
{
var j = theTarget.length;

	for (var i=0; i < theSource.length;i++)
	{
		theTarget.options[j] = new Option(theSource.options[i].text, theSource.options[i].value);
		j++;
	}
	for (var i=0; i < theSource.length;i++)
	{
		theSource.options.remove(i);
		i--;
	}

}

function GetSelectedItems(theSelect)
{
var strSelectedItems = "";

	if (theSelect.length > 0)
	{
	for (var i=0; i < theSelect.length;i++)
	{
		if (theSelect.options[i].selected)
		{
			if (strSelectedItems == "")
			{
				strSelectedItems += theSelect.options[i].value;
			}else{
				strSelectedItems += c_strDelimeter + theSelect.options[i].value;
			}
		}
	}
	return(strSelectedItems);
	}
}

function GetList(theSelect)
{
var strList = "";

	if (theSelect.length > 0)
	{
		for (var i=0; i < theSelect.length;i++)
		{
			if (strList == "")
			{
				strList = theSelect.options[i].value;
			}else{
				strList += c_strDelimeter + theSelect.options[i].value;
			}
		}
	}
	return(strList);
}

function btnDeleteZone_onclick()
{
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + theForm.ZoneName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "DeleteZone";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnReset_onclick()
{

	InitializeLists();
	pblnAddSub = false;

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    document.all("spanZoneName").innerHTML = strDetailTitle;

    theForm.btnUpdate.disabled = false;
    theForm.btnDeleteZone.disabled = false;
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theForm.Action.value = "View";
	theForm.submit();
	return false;
}

function btnSubmit_onclick()
{

if (ValidInput())
{
theForm.ZoneCountry.value = GetList(frmData.targetZoneCountry);
theForm.ZoneState.value = GetList(frmData.targetZoneState);

}else{
	return false;
}

}

function ValidInput()
{
  if (theForm.ZoneName.value == "")
  {
    alert("Please enter a Zone name.")
    theForm.ZoneName.focus();
    return(false);
  }
	
    return(true);
}

//-->
</SCRIPT>

<BODY onload="body_onload();">
<CENTER>
<TABLE border=0 cellPadding=5 cellSpacing=1 width="95%">
  <TR>
    <TH><div class="pagetitle "><%= mstrPageTitle %></div></TH>
  </TR>
</TABLE>

<%= .OutputSummary %>

<FORM action='sszbsZoneAdmin.asp' id=frmData name=frmData method=post>
<input type=hidden id=ZoneID name=ZoneID value=<%= .ZoneID %>>
<input type=hidden id=ZoneCountry name=ZoneCountry value='<%= .ZoneCountries %>'>
<input type=hidden id=ZoneState name=ZoneState value='<%= .ZoneStates %>'>
<input type=hidden id=Action name=Action value='Update'>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanZoneName"><%= .ZoneName %> Details</span></th>
  </tr>
  <tr>
	<th colspan="2" align=center><%= .OutputMessage %></th>
  </tr>
      <TR>
        <TD>Name:</TD>
        <TD><INPUT id=ZoneName name=ZoneName Value='<%= .ZoneName %>' maxlength=30 size=30></TD>
      </TR>
</table>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
	<TR>
        <TD align=center>Available Countries<br />
			<select size=20 id=sourceZoneCountry name=sourceZoneCountry style="width=200;" multiple>
			<% 
				Call MakeCombo("Select loclctryAbbreviation,loclctryName from sfLocalesCountry Order By loclctryName","loclctryName","loclctryAbbreviation","")
			%>
			</select>
        </TD>
        <TD align=center>
			<INPUT class='butn' id=btnAddCountry name=btnAddCountry type=button value='-->' onclick='return MoveItems("Country",true)'><br />
			<INPUT class='butn' id=btnRemoveCountry name=btnRemoveCountry type=button value='<--' onclick='return MoveItems("Country",false)'>
		</TD>        
        <TD align=center>Countries in this Zone<br />
			<select size=20 size="1"  id=targetZoneCountry name=targetZoneCountry style="width=200;" multiple>
			</select>
		</TD>        
        <TD align=center>Available States<br />
			<select size=20 id=sourceZoneState name=sourceZoneState style="width=200;" multiple>
			<% 
				Call MakeCombo("Select loclstAbbreviation,loclstName from sfLocalesState Order By loclstName","loclstName","loclstAbbreviation","")
			%>
			</select>
        </TD>
        <TD align=center>
			<INPUT class='butn' id=btnAddState name=btnAddState type=button value='-->' onclick='return MoveItems("State",true)'><br />
			<INPUT class='butn' id=btnRemoveState name=btnRemoveState type=button value='<--' onclick='return MoveItems("State",false)'>
		</TD>        
        <TD align=center>States in this Zone<br />
			<select size=20 size="1"  id=targetZoneState name=targetZoneState style="width=200;" multiple>
			</select>
		</TD>        
      </TR>
	<TR>
    <TD colspan=2 align=right>
		Postal Codes in this zone<br />(separate by semi-colons):
    </TD>
    <TD colspan=4 align=left>
        <textarea id=ZoneZip name=ZoneZip type=submit style="width=600;"><%= .ZoneZIPs %>
        </textarea>
    </TD>
  </TR>
	<TR>
    <TD colspan=6 align=center>
        <INPUT class='butn' id=btnNewZone name=btnNewZone type=button value='New Zone' onclick='return btnNewZone_onclick()'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick()'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDeleteZone name=btnDeleteZone type=button value='Delete Zone' onclick='return btnDeleteZone_onclick()'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes' onclick='return btnSubmit_onclick()'>
    </TD>
  </TR>
</TABLE>
</FORM>
<!--#include file="zbsadminfooter.asp"-->
</CENTER>
</BODY>
</HTML>
<%
    End With
    Set mclsZone = Nothing
    Set cnn = Nothing
    Response.Flush
%>
