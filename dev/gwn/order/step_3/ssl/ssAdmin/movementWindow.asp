<% Option Explicit 
'********************************************************************************
'*   Promotion Component for StoreFront 5.0										*
'*   Release Version   1.1														*
'*   Release Date      June 2, 2001												*
'*																				*
'*	 Version 1.1 Release Notes													*
'*		- added support for excluding sale items from discount calculations		*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

Response.Buffer = True

'***********************************************************************************************

%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="Common/ssProduct_CommonFilter.asp"-->
<%

'***********************************************************************************************
'***********************************************************************************************

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mstrTable
mstrTable = Request.QueryString("TableSource")
%>
<html>
<head>
<title>Item Mover</title>
<script LANGUAGE=javascript>
<!--
var mobjDic;
var theForm;

function setParentVariables()
{
	theForm = document.frmData;
	//var theTargetDoc = window.parent.opener.document;
	//var theTarget = window.parent.opener;
	mobjDic = eval("window.parent.opener.mdic" + window.parent.opener.gstrstrItemNameTarget);
	FillItem("selTarget");

}

function FillItem(strItemName)
{
	var theTargetSelect = eval("theForm." + strItemName);

	//CleanCategory();
	var pary = (new VBArray(mobjDic.Keys())).toArray();
	var plngKey;
	var theOption;
	
	theTargetSelect.length = 0;
	
	try
	{
	for (var i=0; i < pary.length;i++)
	{
		plngKey = pary[i];
		theOption = new Option(mobjDic(plngKey), plngKey);
		theTargetSelect.options.add(theOption);
	}
	}
	catch(e)
	{
	return false;
	}
}

function MoveItem(strItemNameSource, strItemNameTarget, blnLeft)
{

	var theSourceSelect = eval("theForm." + strItemNameSource);
	var theTargetSelect = eval("theForm." + strItemNameTarget);
	
	var mblnAdded = false;
	
	if (blnLeft)
	{
		if (theSourceSelect.length > 0)
		{
			for (var i=0; i < theSourceSelect.length;i++)
			{
				if (theSourceSelect.options[i].selected)
				{
					if (!mobjDic.Exists(theSourceSelect.options[i].value))
					{
					mblnAdded = true;
					mobjDic.Add (theSourceSelect.options[i].value,theSourceSelect.options[i].text)
					}
				}
			}
		}
	}else{
		for (var i=theTargetSelect.length-1; i >=0 ;i--)
		{
			if (theTargetSelect.options[i].selected)
			{
				mobjDic.Remove(theTargetSelect.options[i].value);
				mblnAdded = true;
			}
		}
	}
	if (mblnAdded)
	{
		FillItem(strItemNameTarget);
		window.parent.opener.FillItem(window.parent.opener.gstrstrItemNameTarget);
	}
}

//-->
</script>
</head>
<body onload="setParentVariables();">
<CENTER>

<form name="frmData" id="frmData" action="" onsubmit="return false;">
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  
  <tr>
    <td>&nbsp;</td>
    <td>
      <table border="0" ID="Table1">
		<tr>
		  <th>Source</th>
		  <th>&nbsp;</th>
		  <th>Target</th>
		</tr>
		<tr>
		  <td align=center>
			<select id="selSource" name="selSource" size="5" ondblclick="MoveItem('selSource','selTarget',true);" multiple>
			<%= MakeCombo_Saved(mstrTable, "") %>
			</select>
		  </td>
		  <td valign=middle align=center>
			<input class="butn" type=button id="btnAddItem" name="btnAddItem" onclick="MoveItem('selSource','selTarget',true);" value="-->"><br />
			<input class="butn" type=button id="btnDeleteItem" name="btnDeleteItem" onclick="MoveItem('selSource','selTarget',false);" value="<--"><br />
		  </td>
		  <td align=center>
			<select id="selTarget" name="selTarget" size=5 ondblclick="MoveItem('selSource','selTarget',false);" multiple>
			</select>
		  </td>
		</tr>
	  </table>
    </td>
  </tr>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
		<input type=submit class='butn' name=btnClose ID="btnClose" value="Close" onclick="window.close();"> 
	</TD>
  </TR>
</TABLE>
</FORM>
</CENTER>
</body>
</HTML>
<% 

	set cnn = Nothing

Response.Flush
%>
