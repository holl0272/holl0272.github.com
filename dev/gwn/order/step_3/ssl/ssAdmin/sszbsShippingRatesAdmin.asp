<%Option Explicit
'********************************************************************************
'*   Zone Based Shipping					                                    *
'*   Release Version:   2.0														*
'*   Release Date:		January 1, 2003											*
'*   Revision Date:		January 1, 2003											*
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsssShippingRates
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private plngShipRateID
Private plngShipMethod
Private plngShipZone
Private pdblShipRate
Private pdblShipWeight
Private pblnShipRatePercentage

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


Public Property Get ShipRateID()
    ShipRateID = plngShipRateID
End Property

Public Property Get ShipMethod()
    ShipMethod = plngShipMethod
End Property

Public Property Get ShipZone()
    ShipZone = plngShipZone
End Property

Public Property Get ShipRate()
    ShipRate = pdblShipRate
End Property

Public Property Get ShipWeight()
    ShipWeight = pdblShipWeight
End Property

Public Property Get ShipRatePercentage()
    ShipRatePercentage = pblnShipRatePercentage
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

	If rs.EOF Then Exit Sub
    plngShipRateID = trim(rs("ShipRateID"))
    plngShipMethod = trim(rs("ShipMethod"))
    plngShipZone = trim(rs("ShipZone"))
    pdblShipRate = trim(rs("ShipRate"))
    pdblShipWeight = trim(rs("ShipWeight"))
    pblnShipRatePercentage = trim(rs("ShipRatePercentage"))

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        plngShipRateID = Trim(.Item("ShipRateID"))
        plngShipMethod = Trim(.Item("ShipMethod"))
        plngShipZone = Trim(.Item("ShipZone"))
        pdblShipRate = Trim(.Item("ShipRate"))
        pdblShipWeight = Trim(.Item("ShipWeight"))
        pblnShipRatePercentage = CBool(Len(Trim(.Item("ShipRatePercentage"))) > 0)
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "ShipRateID=" & lngID
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues (pRS)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function LoadAll()

'On Error Resume Next

    Set pRS = GetRS("Select * from ssShippingRates Order By ShipWeight, ShipZone")
    If Not (pRS.EOF Or pRS.BOF) Then
    	'If len(plngShipRateID) > 0 Then	Call LoadValues(pRS)
    	Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function DuplicateRateTable(lngSourceID, lngTargeID, blnCopyRates)

dim sql
dim p_rs
dim i

On Error Resume Next

	sql = "Select * from ssShippingRates where ShipMethod=" & lngSourceID
    Set p_rs = GetRS(sql)
	For i=1 to p_rs.RecordCount
		If blnCopyRates Then
			sql = "Insert Into ssShippingRates (ShipMethod,ShipZone,ShipWeight,ShipRate) " _
				& "Values (" _
				& lngTargeID & "," _
				& p_rs("ShipZone").value & "," _
				& p_rs("ShipWeight").value & "," _
				& p_rs("ShipRate").value & ")"
		Else
			sql = "Insert Into ssShippingRates (ShipMethod,ShipZone,ShipWeight,ShipRate) " _
				& "Values (" _
				& lngTargeID & "," _
				& p_rs("ShipZone").value & "," _
				& p_rs("ShipWeight").value & "," _
				& 0 & ")"
		End If
		cnn.Execute sql,,128
		p_rs.MoveNext
	Next
	p_rs.Close
	Set p_rs = Nothing

    If Err.Number = 0 Then
		pstrMessage = "The shipping table was successfully duplicated."
    Else
		pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
	End If
            
End Function	'DuplicateRateTable

'***********************************************************************************************

Public Function Delete(lngShipRateID)

Dim sql

'On Error Resume Next

    sql = "Delete from ssShippingRates where ShipRateID = " & lngShipRateID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Shipping Charge successfully deleted."
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

On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If strErrorMessage = "True" Then
        If Len(plngShipRateID) = 0 Then plngShipRateID = 0

        sql = "Select * from ssShippingRates where ShipRateID = " & plngShipRateID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("ShipMethod") = plngShipMethod
        rs("ShipZone") = plngShipZone
        rs("ShipRate") = pdblShipRate
        rs("ShipWeight") = pdblShipWeight
        rs("ShipRatePercentage") = ABS(pblnShipRatePercentage * -1)

        rs.Update

		pblnError = (Err.number <> 0)
        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>You have already created this shipping zone.<br />Please enter a different shipping method, zone, or weight.</H4><br />"
            Else
				pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
            End If
        ElseIf Err.Number = -2147217873 Then
            If Err.Description = "[Microsoft][ODBC Microsoft Access Driver]Error in row " Then
                pstrMessage = "<H4>You have already created this shipping zone.<br />Please enter a different shipping method, zone, or weight.</H4><br />"
            Else
				pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
            End If
        ElseIf Err.Number <> 0 Then
            pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
        
        plngShipRateID = rs("ShipRateID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The shipping charge was successfully added."
                mlngID = plngShipRateID
            Else
                pstrMessage = "The shipping charge was successfully updated."
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

Dim i,j
Dim pstrURL
 
Dim p_rsShipZones, p_lngNumZones
Dim p_rsZones
Dim p_rsShipWeights, p_lngNumWeights

Const p_TableWidth = 100
Const p_TableHeight = 200

Const p_WeightWidth = 8
Const p_RateMinWidth = 10
Dim p_RateWidth
Dim p_blnShipRatePercentage
Dim p_strShipRateOutput

	'Create Filter, if no filter use first record
	If len(mlngShipMethodFilter) = 0 or mlngShipMethodFilter="0" Then
		If prs.RecordCount > 0 Then
			pRS.MoveFirst
			mlngShipMethodFilter = pRS("ShipMethod")
			pRS.Filter = "ShipMethod=" & mlngShipMethodFilter
		End If
	Else
		pRS.Filter = "ShipMethod=" & mlngShipMethodFilter
	End If
	
	If len(plngShipRateID) = 0 Then
		Call LoadValues(pRS)
	End If
  
	'Create Rows and Columns
	p_lngNumZones = 0
	p_lngNumWeights = 0
	If len(mlngShipMethodFilter) > 0 Then
		set p_rsZones = GetRS("Select ZoneID,ZoneName from ssShipZones Order By ZoneName")
		set p_rsShipZones = GetRS("Select Distinct shipZone from ssShippingRates where shipMethod=" & mlngShipMethodFilter & " Order By shipZone")
		set p_rsShipWeights = GetRS("Select Distinct shipWeight from ssShippingRates where shipMethod=" & mlngShipMethodFilter & " Order By shipWeight")
		p_lngNumZones = p_rsShipZones.RecordCount
		p_lngNumWeights = p_rsShipWeights.RecordCount

		If p_lngNumZones > 0 Then
			p_RateWidth = (100 - p_WeightWidth) / p_lngNumZones
		Else
			p_RateWidth = (100 - p_WeightWidth)
		End If
		
'		If p_RateWidth < p_RateMinWidth Then p_RateWidth = p_RateMinWidth
	End If
  
	'Create Header		rules='none' 
    Response.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' rules='none'>" & vbcrlf
	If p_lngNumZones > 0 Then
		Response.Write "	<tr class='tblhdr'>" & vbcrlf
		Response.Write "	  <th colspan='" & (p_lngNumZones + 1) & "'>" & vbcrlf
		Response.Write "<form action='sszbsShippingRatesAdmin.asp' method='post' id=form1 name=form1>"  & vbcrlf
		Response.Write "<input type=hidden id=Action name=Action value='Filter'>"  & vbcrlf
		Response.Write "<SELECT id=ShipMethodFilter name=ShipMethodFilter onchange='this.form.submit();'>"  & vbcrlf
        Call MakeCombo(mrsShippingMethods,"shipMethod","shipID",mlngShipMethodFilter)
		Response.Write "	</SELECT></form>"
		Response.Write "</form>" & vbcrlf
		Response.Write "	</th></tr>" & vbcrlf
		If p_lngNumZones > 0 Then
			Response.Write "	<tr class='tblhdr'><th>&nbsp;</th><th align=center colspan='" & p_lngNumZones & "'>Zone</th></tr>" & vbcrlf
			Response.Write "	<tr class='tblhdr'>" & vbcrlf
			Response.Write "	<th align=center width='" & p_WeightWidth - 1 & "%'>&nbsp;&nbsp;&nbsp;Units</th>" & vbcrlf
			For i=1 to p_lngNumZones
				p_rsZones.Filter = "ZoneID=" & p_rsShipZones("shipZone")
				If 	p_rsZones.EOF Then
					pstrURL = "<a class='HeaderLink' href='ssZoneAdmin.asp?Action=View&ZoneID=" & p_rsShipZones("shipZone") & "' title='Configure Zone'>" & p_rsShipZones("shipZone") & "</a>"
				Else
					pstrURL = "<a class='HeaderLink' href='ssZoneAdmin.asp?Action=View&ZoneID=" & p_rsShipZones("shipZone") & "' title='Configure " & p_rsZones("ZoneName") & "'>" & p_rsZones("ZoneName") & "</a>"
				End If
				Response.Write "	<th align=center width='" & p_RateWidth & "%'>" & pstrURL & "</th>" & vbcrlf
				p_rsShipZones.MoveNext
			Next
			p_rsShipZones.MoveFirst
			Response.Write "	</tr>" & vbcrlf
			Response.Write "<tr><td colspan='" & (p_lngNumZones + 1) & "'>" & vbcrlf
		End If
	Else
		Response.Write "	<tr class='tblhdr' colspan='" & p_lngNumZones & "'>" & vbcrlf
		Response.Write "	  <th>" & vbcrlf
		Response.Write "<form action='sszbsShippingRatesAdmin.asp' method='post' id=form1 name=form1>"  & vbcrlf
		Response.Write "<input type=hidden id=Action name=Action value='Filter'>"  & vbcrlf
		Response.Write "<SELECT id=ShipMethodFilter name=ShipMethodFilter onchange='this.form.submit();'>"  & vbcrlf
        Call MakeCombo("Select shipID,shipMethod from sfShipping Order By shipMethod","shipMethod","shipID",mlngShipMethodFilter)
		Response.Write "	</SELECT></form>"
		Response.Write "</form>" & vbcrlf
		Response.Write "	</th></tr>" & vbcrlf
		Response.Write "<tr><td colspan='" & p_lngNumZones & "'>" & vbcrlf
	End If
	
    Response.Write "<div name='divSummary' style='height:" & p_TableHeight & "; overflow:auto;'>" & vbcrlf
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary' id='tblSummary' " _
				 & ">" & vbcrlf


    Response.Write "<COLGROUP align='left'>" & vbcrlf
	For i=1 to p_lngNumZones
		Response.Write "<COLGROUP align='center'>" & vbcrlf
	Next

	If prs.RecordCount = 0 Then
        Response.Write "<TR><TD align='center'><h3>There are no Shipping Rates for this Shipping Method</h3></TD></TR>" & vbcrlf
    End If

	For j=1 to p_lngNumWeights
		Response.Write "<TR><TH class='tblhdr' align=center width='" & p_WeightWidth & "%'>" & p_rsShipWeights("shipWeight") & "</TH>" & vbcrlf
		p_rsShipZones.MoveFirst
		For i=1 to p_lngNumZones
			prs.Filter = "ShipMethod=" & mlngShipMethodFilter _
					   & " AND shipWeight=" & p_rsShipWeights("shipWeight") _
					   & " AND shipZone=" & p_rsShipZones("shipZone")

			If prs.EOF Then
			    Response.Write " <TD title='Click to create this rate' width='" & p_RateWidth & "%'" _
							 & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
							 & "onmouseout='doMouseOutRow(this); ClearTitle();' " _
							 & "onmousedown=" & chr(34) & "FillIn(" & p_rsShipWeights("shipWeight") & "," & p_rsShipZones("shipZone") & ")" & chr(34) & ">" _
							 & "-" & "</TD>" & vbcrlf
			Else
				pstrURL = "sszbsShippingRatesAdmin.asp?Action=View&ShipRateID=" & prs("ShipRateID")
				
				If Len(prs.Fields("ShipRatePercentage").Value & "") = 0 Then
					p_blnShipRatePercentage = False
				ElseIf CBool(prs.Fields("ShipRatePercentage").Value) Then
					p_blnShipRatePercentage = True
				Else
					p_blnShipRatePercentage = False
				End If
				
				If p_blnShipRatePercentage Then
					p_strShipRateOutput = prs.Fields("ShipRate").Value & "%"
				Else
					p_strShipRateOutput = FormatCurrency(prs.Fields("ShipRate").Value,2)
				End If
				
				If trim(prs("ShipRateID")) = plngShipRateID Then
					Response.Write "<TD class='Selected' width='" & p_RateWidth & "%'>" & p_strShipRateOutput & "</TD>" & vbcrlf
				Else
				    Response.Write " <TD title='Click to edit this entry' width='" & p_RateWidth & "%'" _
								 & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
								 & "onmouseout='doMouseOutRow(this); ClearTitle();' " _
								 & "onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">" _
								 & p_strShipRateOutput & "</TD>" & vbcrlf
				End If
			End If
			p_rsShipZones.MoveNext
		Next
		Response.Write "</TR>" & vbcrlf
		p_rsShipWeights.MoveNext
	Next
	
    Response.Write "</td></tr></TABLE></div>"
    Response.Write "</TABLE>"

	set p_rsZones = Nothing
	set p_rsShipZones = Nothing
	set p_rsShipWeights = Nothing

End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Not IsNumeric(plngShipMethod) And Len(plngShipMethod) <> 0 Then
        strError = strError & "Please enter a number for the Shipping Method." & cstrDelimeter
    ElseIf Len(plngShipMethod) = 0 Then
        strError = strError & "Please enter a value for the Shipping Method." & cstrDelimeter
    End If

    If Not IsNumeric(plngShipZone) And Len(plngShipZone) <> 0 Then
        strError = strError & "Please enter a number for the Shipping Zone." & cstrDelimeter
    ElseIf Len(plngShipZone) = 0 Then
        strError = strError & "Please enter a value for the Shipping Zone." & cstrDelimeter
    End If

    If Not IsNumeric(pdblShipRate) And Len(pdblShipRate) <> 0 Then
        strError = strError & "Please enter a number for the Shipping Charge." & cstrDelimeter
    ElseIf Len(pdblShipRate) = 0 Then
        strError = strError & "Please enter a value for the Shipping Charge." & cstrDelimeter
    End If

    If Not IsNumeric(pdblShipWeight) And Len(pdblShipWeight) <> 0 Then
        strError = strError & "Please enter a number for the Order Weight." & cstrDelimeter
    ElseIf Len(pdblShipWeight) = 0 Then
        strError = strError & "Please enter a value for the Order Weight." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)

End Function 'ValidateValues

End Class   'clsssShippingRates
'--------------------------------------------------------------------------------------------------
%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'**************************************************
'
'	Start Code Execution
'

Dim mrsShippingMethods

Dim mAction
Dim mclsssShippingRates
Dim mlngID, mlngShipMethodFilter
Private mstrShipMethodName

	'call InitializeConnection(cnn)
	mstrPageTitle = "Weight Based Shipping Charge Administration"


    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    
    mlngID = Request.QueryString("ShipRateID")
    If Len(mlngID) = 0 Then mlngID = Request.Form("ShipRateID")
    
    mlngShipMethodFilter = Request.QueryString("ShipMethodFilter")
    If Len(mlngShipMethodFilter) = 0 Then mlngShipMethodFilter = Request.Form("ShipMethodFilter")

    Set mclsssShippingRates = New clsssShippingRates

    Select Case mAction
        Case "Update"
            mclsssShippingRates.Update
            If mclsssShippingRates.LoadAll Then mclsssShippingRates.Find mlngID
        Case "New"
            mclsssShippingRates.Update
            If mclsssShippingRates.LoadAll Then mclsssShippingRates.Find mlngID
        Case "DuplicateRateTable"
			mlngShipMethodFilter = Request.Form("ShipMethod")
            mclsssShippingRates.DuplicateRateTable Request.Form("origShipMethod"),Request.Form("ShipMethod"),cbool(Request.Form("CopyRates"))
            mclsssShippingRates.LoadAll
        Case "Delete"
            mclsssShippingRates.Delete mlngID
            mclsssShippingRates.LoadAll
        Case "View"
            If mclsssShippingRates.LoadAll Then 
				mclsssShippingRates.Find mlngID
				mlngShipMethodFilter = mclsssShippingRates.ShipMethod
			End If
        Case "Filter"
            mclsssShippingRates.LoadAll
        Case Else
            If mclsssShippingRates.LoadAll Then 
				mlngID = mclsssShippingRates.ShipRateID
				mlngShipMethodFilter = mclsssShippingRates.ShipMethod
			End If
    End Select
    
    Set mrsShippingMethods = GetRS("Select shipID,shipMethod from sfShipping Order By shipMethod")
    If mrsShippingMethods.RecordCount > 0 Then
		If len(mlngShipMethodFilter) > 0 Then
			mrsShippingMethods.Filter = "shipID=" & mlngShipMethodFilter
			If not mrsShippingMethods.EOF Then mstrShipMethodName = mrsShippingMethods("shipMethod").value
			mrsShippingMethods.filter = ""
		End If
    End If
    
Call WriteHeader("",True)
    With mclsssShippingRates
%>


<SCRIPT LANGUAGE=javascript>
<!--

function HighlightValue(theSelect,theValue)
{
  for (var i = 0;  i < theSelect.options.length;  i++)
  {
	if (theSelect.options[i].value == theValue)
	{
	theSelect.selectedIndex = i;
	return true;
	}
  }
}

function FillIn(dblWeight,lngZone)
{
var theForm = document.frmData;

	btnNew_onclick(theForm.btnUpdate);
    theForm.ShipWeight.value = dblWeight;
    HighlightValue(theForm.ShipZone,lngZone)
    theForm.ShipRate.focus();
    theForm.ShipRate.select();
	
return(true);
}

function SetDefaults(theForm)
{
    theForm.ShipRateID.value = "";
    theForm.ShipRate.value = "0";
    theForm.ShipWeight.value = "0";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.Action.value = "New";
    theForm.btnUpdate.value = "Add New Shipping Rate";
    theForm.btnDelete.disabled = true;
	theForm.ShipZone.focus();
//	theForm.ShipZone.select();
}

function DuplicateRateTable(theButton)
{
var theForm = theButton.form;

var strTarget = theForm.ShipMethod.value;
var strSource = theForm.origShipMethod.value;
var blnConfirm;
var blnSetRates;

	if (strTarget == strSource)
	{
		alert("You cannot copy the rate table onto iteself. \n Please select a different rate table to duplicate this to.");
		theForm.ShipMethod.focus();
		return false;
	}else{
		blnConfirm = confirm("Are you sure you want to duplicate this rate table?");
		if (blnConfirm)
		{
			blnSetRates = confirm("Select OK to copy the rate table structure and rates. \n Select Cancel to copy the rate table structure only.");
			theForm.Action.value = "DuplicateRateTable";
			theForm.CopyRates.value = blnSetRates;
			theForm.submit();
			return true;
		}
	}
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete this shipping charge?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }else{
    return(false);
    }
}

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
}

function ValidateInput(theForm)
{
  if (!isNumeric(theForm.ShipZone,false,"Please enter a value for the shipping zone.")) {return(false);}
  if (!isNumeric(theForm.ShipWeight,false,"Please enter a value for the shipping weight.")) {return(false);}
  if (!isNumeric(theForm.ShipRate,false,"Please enter a value for the shipping charge.")) {return(false);}
  
  return(true);
}

//-->
</SCRIPT>
<CENTER>
<body>

<p><div class="pagetitle "><%= mstrPageTitle %></div></p>

<%= .OutputSummary %>

<FORM action='sszbsShippingRatesAdmin.asp' id=frmData name=frmData onsubmit='return ValidateInput(this);' method=post>
<input type=hidden id=ShipRateID name=ShipRateID value=<%= .ShipRateID %>>
<input type=hidden id=origShipMethod name=origShipMethod value=<%= .ShipMethod %>>
<input type=hidden id=CopyRates name=CopyRates value="">
<input type=hidden id=ShipMethodFilter name=ShipMethodFilter value=<%= mlngShipMethodFilter %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit Value Based Shipping Entry</span></th>
  </tr>
  <tr>
	<td colspan="2" align=center><%= .OutputMessage %></td>
  </tr>

      <TR>
        <TD class="Label">Shipping Method:</TD>
        <TD>
			<SELECT id='ShipMethod' name='ShipMethod'>
            <% Call MakeCombo(mrsShippingMethods,"shipMethod","shipID",.ShipMethod) %>
			</SELECT>&nbsp;&nbsp;
			<% If len(mstrShipMethodName) > 0 Then Response.Write "<a href='ssCarrierAdmin.asp?Action=View&ID=" & .ShipMethod & "'>Configure " & mstrShipMethodName & "</a>" %>
		</TD>
      </TR>
      <TR>
        <TD class="Label">Shipping Zone:</TD>
        <TD>
			<SELECT id='ShipZone' name='ShipZone'>
            <% Call MakeCombo("Select ZoneID,ZoneName from ssShipZones Order By ZoneName","ZoneName","ZoneID",.ShipZone) %>
			</SELECT>&nbsp;&nbsp;
		</TD>
      </TR>
      <TR>
        <TD class="Label">Units (item count, subtotal, weight)  up to:</TD>
        <TD><INPUT id='ShipWeight' name='ShipWeight' Value='<%= .ShipWeight %>'></TD>
      </TR>
      <TR>
        <TD class="Label">Incur a shipping charge of:</TD>
        <TD>
          <INPUT id='ShipRate' name='ShipRate' Value='<%= .ShipRate %>'>&nbsp;
          <INPUT id=ShipRatePercentage name=ShipRatePercentage type=checkbox value="1" <% if .ShipRatePercentage then Response.Write "checked" %>><label id=lblShipRatePercentage for=ShipRatePercentage>Check if Percentage</label>
        </TD>
      </TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick(this)'>&nbsp;
        <INPUT class='butn' id=btnDuplicate name=btnDuplicate type=button value='Copy Rate Table' onclick='return DuplicateRateTable(this);'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
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
    Set mclsssShippingRates = Nothing
    Set mrsShippingMethods = Nothing
    Set cnn = Nothing
    Response.Flush
%>
