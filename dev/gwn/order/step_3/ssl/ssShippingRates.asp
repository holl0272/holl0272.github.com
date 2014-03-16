<%
Option Explicit
'********************************************************************************
'*   Postage Rate Component	for StoreFront 2000/5.0								*
'*   Release Version   2.0.8													*
'*   Release Date      October 27, 2002											*
'*   Revision Date     March 1, 2003											*
'*																				*
'*   Release Notes:                                                             *
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************
%>
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incgeneral.asp"-->
<!--#include file="SFLib/incae.asp"-->
<%
'**********************************************************
'*	Functions
'**********************************************************

	Sub cleanupPageObjects

	On Error Resume Next

		Set mclsCartTotal = Nothing
		Call cleanup_dbconnopen

	End Sub

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Shipping Rate Estimator</title>
</head>
<script language="javascript" type="text/javascript">

var theTargetForm;
var theShippingDropdown;
var theShippingTextbox;
var theShippingHidden;

function setTargetForm()
{
	var theTargetDoc = window.parent.opener.document;
	var bytTarget = theTargetDoc.forms.length - 1;
	theTargetForm = theTargetDoc.forms[bytTarget];
	//alert(theTargetForm.name);
	
	<% If cblnSF5AE Then %>
	theShippingDropdown = theTargetForm.Shipping;
	theShippingTextbox = theTargetForm.txtDummyShip;
	theShippingHidden = theTargetForm.Shipping;
	<% Else %>
	theShippingDropdown = theTargetForm.ShipMeth;
	<% End If %>

}

function MakeSelection_required(strKey,strText,strOther)
{
	setTargetForm();
	theShippingHidden.value = strKey;
	theShippingTextbox.value = strText;
	
	window.close();
	
return true;
}

function MakeSelection(strKey,strText,strOther)
{

  setTargetForm();
  if (strOther != ""){theTargetForm.ShippingSelection.value = strOther};
  for (var i = 0;  i < theShippingDropdown.options.length;  i++)
  {
  if (theShippingDropdown.options[i].value == strKey)
//  if (theShippingDropdown.options[i].text == strKey)
  {
  theShippingDropdown.selectedIndex = i;
  }
  }
window.close();
return true;
}

function MakeSelectionRadio(theForm)
{
var blnFound = false;
var strKey;

	setTargetForm();

	if (theForm.radio1.length > 0)
	{
	  for (var i = 0;  i < theForm.radio1.length;  i++)
	  {
		if (theForm.radio1[i].checked)
		{
		strKey = theForm.radio1[i].value;
		blnFound = true;
		}
	  }
	}else{
		blnFound = theForm.radio1.checked
		strKey = theForm.radio1.value;
	}
	
  if (!blnFound)
  {
  alert("Please select a shipping method.");
  return false;
  }

  for (var i = 0;  i < theShippingDropdown.options.length;  i++)
  {
	if (theShippingDropdown.options[i].value == strKey)
	{
	theShippingDropdown.selectedIndex = i;
	}
  }

window.close();
return true;
}

function clearMessage()
{
var objWaitMessage;

	if (ns4)
		objWaitMessage=document.waitMessage;
	else if (ns6)
		objWaitMessage=document.getElementById("waitMessage").style;
	else if (ie4)
		objWaitMessage=document.all.waitMessage.style;

	if(ns4)
	{
		objWaitMessage.visibility="hidden";
	}
	else if (ns6||ie4)
	{
		objWaitMessage.display = "none";
	}

}

var mstrWaitMessage = '<center><h4>Please wait while we obtain your rates . . .</h4></center>';

var ns4=document.layers;
var ns6=document.getElementById&&!document.all;
var ie4=document.all;

if (ns6||ie4){
	document.write('<div id="waitMessage" style="display:">' + mstrWaitMessage + '</div>')
}
else if (ns4){
	document.write('<layer name="waitMessage" visibility="visible"><div id="waitMessage" style="display:">' + mstrWaitMessage + '</div></layer>')
}

</script>
<body onload="clearMessage();">
<%
'****************************************************************************************************************

'Send the wait message to the browser
If Response.Buffer Then Response.Flush

Dim mclsShipping
Dim mstrOriginStateAbb
Dim mstrOriginZip
Dim mstrOriginCountryAbb

Dim mstrDestinationStateAbb
Dim mstrDestinationZip
Dim mstrDestinationCountryAbb
Dim mblnOrderLoaded
Dim mblnShowFreightQuoteRates
Dim mdblsubTotal

	Call InitializeCart
	With mclsCartTotal

		.City = visitorCity
		.State = visitorState
		.ZIP = visitorZIP
		.Country = visitorCountry
		.isCODOrder = False

		.ShipMethodCode = visitorPreferredShippingCode
		.LoadAllShippingMethods = False
		mblnOrderLoaded = LoadOrderItems_SF5(.OrderItems)
		mdblsubTotal = .subTotal
		
		'.checkInventoryLevels
		'.writeDebugCart

'		'.displayOrder
	End With	'mclsCartTotal
	
	If mblnOrderLoaded Then
		Call InitializeOrigin(mstrOriginStateAbb, mstrOriginZip, mstrOriginCountryAbb)
				
		mstrDestinationStateAbb = Request.QueryString("DestinationState")
		mstrDestinationZip = Request.QueryString("DestinationZip")
		mstrDestinationCountryAbb = Request.QueryString("DestinationCountry")

		mblnShowFreightQuoteRates = Request.QueryString("ShowFreightQuoteRates")
		If Len(mblnShowFreightQuoteRates) > 0 Then
			mblnShowFreightQuoteRates = CBool(mblnShowFreightQuoteRates)
		Else
			mblnShowFreightQuoteRates = False
		End If

		Set mclsShipping = New clsShipping
		With mclsShipping

			.Connection = cnn
			.OriginStateAbb = mstrOriginStateAbb
			.OriginZip = mstrOriginZip
			.OriginCountryAbb = mstrOriginCountryAbb
			
			.DestinationStateAbb = mstrDestinationStateAbb
			.DestinationZIP = mstrDestinationZip
			.DestinationCountryAbb = mstrDestinationCountryAbb
			.DestinationCountryName = GetDestinationCountryName(mstrDestinationCountryAbb)

			.ShowFreightQuoteRates = mblnShowFreightQuoteRates
			
			'Check if Residential Delivery is set
			If Len(Request.QueryString("ShipResidential")) > 0 Then
				.ResidentialDelivery = CBool(Request.QueryString("ShipResidential"))
			End If
			
			'Check if Indoor Delivery is set
			If Len(Request.QueryString("InsideDelivery")) > 0 Then
				.InsideDelivery = CBool(Request.QueryString("InsideDelivery"))
			End If
			
			.OrderSubtotal = mdblsubTotal
			.DeclaredValue = mdblsubTotal
			.Insured = False

			.OrderItems = maryOrderItems
			.MaxItemWeight = mdblMaxItemWeight
			.TotalOrderWeight = mdblOrderWeight
			
			.GetRates ""
		End With
	End If

	If mblnOrderLoaded Then
		Select Case Trim(Request.QueryString("DisplayStyle"))
			Case 0: mclsShipping.DisplayShippingOptions False
			Case 1: mclsShipping.DisplayShippingOptionsAsRadio
			Case 2: mclsShipping.DisplayShippingOptions True
		End Select
		If Not mblnShowFreightQuoteRates And mclsShipping.FreightQuoteEnabled Then
%>
	<tr>
		<td align='middle' colspan='2'>Click <a href="ssShippingRates.asp?<%= Request.QueryString %>&ShowFreightQuoteRates=True">here</a> to obtain rates for shipment via truck</td>
	</tr>
<%
		End If	'mblnShowFreightQuoteRates
	Else
%>
<center>
<h3>It appears your shopping session has expired</h3>
<p><a href='' onclick='window.close(); return false;'>Close window</a></p>
</center>
<%
	End If

set mclsShipping = Nothing
Call cleanupPageObjects
%>
</body>
</html>