<%Option Explicit
'********************************************************************************
'*   Postage Rate Administration						                        *
'*   Release Version: 2.0			                                            *
'*   Release Date: September 21, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsShipMethod
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsProducts
Private pblnError

'database variables

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsProducts)
End Sub

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

Public Property Get rsProducts()
    Set rsProducts = prsProducts
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

Public Function Load()

dim pstrSQL
dim p_strWhere
dim i
dim sql

	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		pstrSQL = "SELECT * FROM ssShippingMethods ORDER BY ssShippingMethodEnabled, ssShippingMethodOrderBy, ssShippingMethodName"
		'debugprint "pstrSQL",pstrSQL
		'Response.Flush	  
		  
		On Error Resume Next
		If Err.number <> 0 Then Err.Clear
		
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			Response.Write "<h3><font color=red>The Postage Rate add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
			Response.Write "<a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PostageRate'>Click here to upgrade</a></h3>"
			Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
			Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
			Response.Flush
			Err.Clear
			Load = False
			Exit Function
		End If
		On Error Goto 0
		
	End With

    Load = (Not prsProducts.EOF)

End Function    'Load

'***********************************************************************************************

Public Function Add()

Dim sql

	sql = "Insert Into ssShippingMethods (ssShippingCarrierID,ssShippingMethodCode,ssShippingMethodName,ssShippingMethodEnabled,ssShippingMethodIsSpecial,ssShippingMethodLocked,ssShippingMethodMinCharge,ssShippingMethodMultiple,ssShippingMethodPerPackageFee,ssShippingMethodPerShipmentFee,ssShippingMethodOfferFreeShippingAbove,ssShippingMethodLimitFreeShippingByWeight,ssShippingMethodClass,ssShippingMethodOrderBy,ssShippingMethodDefault,ssShippingMethodMinWeight,ssShippingMethodPrefWeight,ssShippingMethodMaxLength,ssShippingMethodMaxWidth,ssShippingMethodMaxHeight,ssShippingMethodMaxWeight,ssShippingMethodMaxGirth)" _
		& " Values (1,'shipCode','shipName',0,0,0,0,1,0,0,999999,0,0,999,0,0,70,0,0,0,70,170)"		
	'Response.Write "sql: " & sql & "<br />"
	cnn.Execute sql,,128

End Function	'Add

'***********************************************************************************************

Public Function Delete()

Dim sql
Dim paryDirty
Dim i
Dim ssShippingMethodID

	paryDirty = Split(Request.Form("dirty"),",")
	
	'Update the methods
	For i = 0 To UBound(paryDirty)
		ssShippingMethodID = Trim(paryDirty(i))
		sql = "Delete From ssShippingMethods Where ssShippingMethodID=" & ssShippingMethodID
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
		
	Next 'i

End Function	'Add

'***********************************************************************************************

Public Function Update()

Dim sql
Dim strErrorMessage
Dim vItem
Dim paryIDs
Dim paryDirty
Dim i

'On Error Resume Next

    pblnError = False

	paryDirty = Split(Request.Form("dirty"),",")
	
	'For Each vItem in Request.Form
	'	debugprint vItem, Request.Form(vItem)
	'Next

	Dim ssShippingMethodID
	Dim ssShippingCarrierID
	Dim ssShippingMethodCode
	Dim ssShippingMethodName
	Dim ssShippingMethodEnabled
	Dim ssShippingMethodIsSpecial
	Dim ssShippingMethodLocked
	Dim ssShippingMethodMinCharge
	Dim ssShippingMethodMultiple
	Dim ssShippingMethodPerPackageFee
	Dim ssShippingMethodPerShipmentFee
	Dim ssShippingMethodOfferFreeShippingAbove
	Dim ssShippingMethodLimitFreeShippingByWeight
	Dim ssShippingMethodClass
	Dim ssShippingMethodOrderBy
	Dim ssShippingMethodDefault
	Dim ssShippingMethodMinWeight
	Dim ssShippingMethodPrefWeight
	Dim ssShippingMethodMaxLength
	Dim ssShippingMethodMaxWidth
	Dim ssShippingMethodMaxHeight
	Dim ssShippingMethodMaxWeight
	Dim ssShippingMethodMaxGirth
	Dim ssShippingMethodCountryRule

	'Update the methods
	For i = 0 To UBound(paryDirty)

		ssShippingMethodID = Trim(paryDirty(i))
		ssShippingCarrierID = Request.Form("ssShippingCarrierID" & ssShippingMethodID)
		ssShippingMethodCode = Request.Form("ssShippingMethodCode" & ssShippingMethodID)
		ssShippingMethodName = Request.Form("ssShippingMethodName" & ssShippingMethodID)
		ssShippingMethodEnabled = Request.Form("ssShippingMethodEnabled" & ssShippingMethodID)
		ssShippingMethodIsSpecial = Request.Form("ssShippingMethodIsSpecial" & ssShippingMethodID)
		ssShippingMethodLocked = Request.Form("ssShippingMethodLocked" & ssShippingMethodID)
		ssShippingMethodMinCharge = Request.Form("ssShippingMethodMinCharge" & ssShippingMethodID)
		ssShippingMethodMultiple = Request.Form("ssShippingMethodMultiple" & ssShippingMethodID)
		ssShippingMethodPerPackageFee = Request.Form("ssShippingMethodPerPackageFee" & ssShippingMethodID)
		ssShippingMethodPerShipmentFee = Request.Form("ssShippingMethodPerShipmentFee" & ssShippingMethodID)
		ssShippingMethodOfferFreeShippingAbove = Request.Form("ssShippingMethodOfferFreeShippingAbove" & ssShippingMethodID)
		ssShippingMethodLimitFreeShippingByWeight = Request.Form("ssShippingMethodLimitFreeShippingByWeight" & ssShippingMethodID)
		ssShippingMethodClass = Request.Form("ssShippingMethodClass" & ssShippingMethodID)
		ssShippingMethodOrderBy = Request.Form("ssShippingMethodOrderBy" & ssShippingMethodID)
		ssShippingMethodMinWeight = Request.Form("ssShippingMethodMinWeight" & ssShippingMethodID)
		ssShippingMethodPrefWeight = Request.Form("ssShippingMethodPrefWeight" & ssShippingMethodID)
		ssShippingMethodMaxLength = Request.Form("ssShippingMethodMaxLength" & ssShippingMethodID)
		ssShippingMethodMaxWidth = Request.Form("ssShippingMethodMaxWidth" & ssShippingMethodID)
		ssShippingMethodMaxHeight = Request.Form("ssShippingMethodMaxHeight" & ssShippingMethodID)
		ssShippingMethodMaxWeight = Request.Form("ssShippingMethodMaxWeight" & ssShippingMethodID)
		ssShippingMethodMaxGirth = Request.Form("ssShippingMethodMaxGirth" & ssShippingMethodID)
		ssShippingMethodCountryRule = Request.Form("ssShippingMethodCountryRule" & ssShippingMethodID)
		
		If Len(ssShippingMethodEnabled) = 0 Then ssShippingMethodEnabled = 0

		If Len(ssShippingMethodLocked) = 0 Then ssShippingMethodLocked = 0
		sql = "Update ssShippingMethods Set " _
			& makeSQLUpdate("ssShippingCarrierID", ssShippingCarrierID, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodCode", ssShippingMethodCode, False, 0) & ", " _
			& makeSQLUpdate("ssShippingMethodName", ssShippingMethodName, False, 0) & ", " _
			& makeSQLUpdate("ssShippingMethodEnabled", ssShippingMethodEnabled, False, enDatatype_boolean) & ", " _
			& makeSQLUpdate("ssShippingMethodIsSpecial", ssShippingMethodIsSpecial, False, enDatatype_boolean) & ", " _
			& makeSQLUpdate("ssShippingMethodLocked", ssShippingMethodLocked, False, enDatatype_boolean) & ", " _
			& makeSQLUpdate("ssShippingMethodMinCharge", ssShippingMethodMinCharge, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMultiple", ssShippingMethodMultiple, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodPerPackageFee", ssShippingMethodPerPackageFee, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodPerShipmentFee", ssShippingMethodPerShipmentFee, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodOfferFreeShippingAbove", ssShippingMethodOfferFreeShippingAbove, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodLimitFreeShippingByWeight", ssShippingMethodLimitFreeShippingByWeight, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodClass", ssShippingMethodClass, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMinWeight", ssShippingMethodMinWeight, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodPrefWeight", ssShippingMethodPrefWeight, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMaxLength", ssShippingMethodMaxLength, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMaxWidth", ssShippingMethodMaxWidth, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMaxHeight", ssShippingMethodMaxHeight, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMaxWeight", ssShippingMethodMaxWeight, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodCountryRule", ssShippingMethodCountryRule, False, 1) & ", " _
			& makeSQLUpdate("ssShippingMethodMaxGirth", ssShippingMethodMaxGirth, False, 1) & " " _
			& " Where  ssShippingMethodID=" & ssShippingMethodID

			'Response.Write i & ": " & sql & "<br />"
			cnn.Execute sql,,128
	Next 'i
		
	'Now set the display order and default
	paryIDs = Split(Request.Form("ShippingMethodID"),",")
	ssShippingMethodDefault = Trim(Request.Form("ssShippingMethodDefault"))
	For i = 0 To UBound(paryIDs)
		ssShippingMethodID = Trim(paryIDs(i))
		If ssShippingMethodDefault = ssShippingMethodID Then
			sql = "Update ssShippingMethods Set " _
				& makeSQLUpdate("ssShippingMethodOrderBy", i, False, 1) & ", ssShippingMethodDefault=1" _
				& " Where  ssShippingMethodID=" & ssShippingMethodID
		Else
			sql = "Update ssShippingMethods Set " _
				& makeSQLUpdate("ssShippingMethodOrderBy", i, False, 1) & ", ssShippingMethodDefault=0" _
				& " Where  ssShippingMethodID=" & ssShippingMethodID
		End If
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
	Next 'i
	
'		& makeSQLUpdate("ssShippingMethodDefault", ssShippingMethodDefault, False, 1) & " " _
	
    Update = (not pblnError)
    
    Application.Contents.Remove("ShippingMethods")
    If Err.number <> 0 Then Err.Clear

End Function    'Update

'***********************************************************************************************

Function ConvertBoolean(vntValue)

	If Len(Trim(vntValue & "")) = 0 Then
		ConvertBoolean = False
	Else
		On Error Resume Next
		ConvertBoolean = cBool(vntValue)
		If Err.number <> 0 Then 
			ConvertBoolean = False
			Err.Clear
		End If
	End If

End Function	'ConvertBoolean

'******************************************************************************************************************************************************************

End Class   'clsShipMethod

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'******************************************************************************************************************************************************************

'******************************************************************************************************************************************************************

mstrPageTitle = "Shipping Methods Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsShipMethod
Dim mobjrsCarriers

	mAction = LoadRequestValue("Action")

    Set mclsShipMethod = New clsShipMethod
    With mclsShipMethod
		If mAction = "Update" Then .Update
		If mAction = "Add" Then .Add
		If mAction = "Delete" Then .Delete
		.Load
	End With
    
	Set mobjrsCarriers = GetRS("Select ssShippingCarrierID, ssShippingCarrierName from ssShippingCarriers Order By ssShippingCarrierName")

	Call WriteHeader("",True)
%>
<SCRIPT LANGUAGE="javascript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--

function CheckAll(blnCheck)
{
	var plngCount;
	var i;

	plngCount = document.frmData.dirty.length;
	if (document.frmData.dirty.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
		document.frmData.dirty[i].checked = blnCheck;
		}
	}else{
	document.frmData.dirty.checked = blnCheck;
	}
	
}

function makeDirty(theItem, lngID)
{

	var plngCount;
	var i;

	plngCount = document.frmData.dirty.length;
	if (document.frmData.dirty.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
			if (document.frmData.dirty[i].value == lngID)
			{
			document.frmData.dirty[i].checked = true;
			return true
			}
		}
	}

}

function moveCarrier(blnUp)
{
	var tblSummary = document.all("tblSummary");
	
	if (blnUp)
	{
		tblSummary.rows[mlngSelectedRow].swapNode(tblSummary.rows[mlngSelectedRow-1]);
		mlngSelectedRow = mlngSelectedRow - 1;
	}else{
		tblSummary.rows[mlngSelectedRow].swapNode(tblSummary.rows[mlngSelectedRow+1]);
		mlngSelectedRow = mlngSelectedRow + 1;
	}
	
	setMoveButtons();
	
}

function setMoveButtons()
{
var tblSummary = document.all("tblSummary");
var plngCount = tblSummary.rows.length - 3;

	document.frmData.btnMoveUp.disabled = (mlngSelectedRow == 2);
	document.frmData.btnMoveDown.disabled = (mlngSelectedRow == plngCount);
}

var mlngSelectedRow;

function highlightRow(theRow, blnHighlight)
{
var tblSummary = document.all("tblSummary");
var plngCount = tblSummary.rows.length - 3;
var i;

	mlngSelectedRow = theRow.rowIndex;
	
	for (i=2; i < tblSummary.rows.length-2;i++)
	{
	tblSummary.rows[i].className = "Inactive";
	}
	if (blnHighlight)
	{
	theRow.className = "Selected";
	}else{
	theRow.className = "Inactive";
	}
	
	setMoveButtons();
	
}

//-->
</SCRIPT>

<CENTER>

<table border=0 cellPadding=5 cellSpacing=1 id="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
</table>

<FORM action="ssPostageRate_shippingMethodsAdmin.asp" id="frmData" name="frmData" method="post">
<input type="hidden" id="Action" name="Action" value="Update">

<table class="tbl" cellpadding="2" cellspacing="0" border="1" id="tblSummary">
  <tr class="tblhdr">
    <th>&nbsp;</th>
    <th colspan="6">Shipping Methods</th>
    <th colspan="4">Shipping Charges</th>
    <th colspan="2">Free Shipping</th>
    <th colspan="2">Order Restrictions</th>
    <th colspan="2">Package Restrictions</th>
    <th colspan="2">&nbsp;</th>
  </tr>
  <tr class="tblhdr">
    <th><div title="Update shipping method">&nbsp;</div></th>
    <th><div title="Record ID">ID</div></th>
    <th><div title="Shipping Method Name - appears in dropdown on process_order.asp and in reports">Name</div></th>
    <th><div title="Shipping Method Code - these are set by the shipping carriers and should not be editied">Code</div></th>
    <th><div title="Shipping carrier">Carrier</div></th>
    <th><div title="Is this a shipping method you wish to support?">Enabled</div></th>
    <th><div title="Is this a shipping method only applicable to specific products?">Special</div></th>
    <th><div title="Minimum shipping charge for this shipping method - set to 0 if you intend to offer free shipping">Charge</div></th>
    <th><div title="the amount to multiply the carriers base rate by">Mult.</div></th>
    <th><div title="amount to charge per unique package">Pkg Fee</div></th>
    <th><div title="amount to charge per order">Ship Fee</div></th>
    <th><div title="amount of order which qualifies for this method to be free">Amount</div></th>
    <th><div title="weight limit of order which qualifies for free shipping">Max Wt</div></th>
    <th><div title="minimum weight of the order before this shipping method is available">Min Wt</div></th>
    <th><div title="maximum weight of order where this shipping method is available">Max Wt</div></th>
    <th><div title="preferred packaging weight">Pref Wt</div></th>
    <th><div title="maximum girth of package">Max Girth</div></th>
    <th><div title="default method to be selected">Default</div></th>
    <th><div title="select countries this method applies to">Zones</div></th>
    <!--
    <th><div title="">ssShippingMethodMaxLength</div></th>
    <th><div title="">ssShippingMethodMaxWidth</div></th>
    <th><div title="">ssShippingMethodMaxHeight</div></th>
    <th><div title="">ssShippingMethodClass</div></th>
    <th><div title="">ssShippingMethodLocked</div></th>
    -->
  </tr>

  <% 
  Dim plngShippingMethodID
  
  With mclsShipMethod.rsProducts 
	Do While Not .EOF
		plngShippingMethodID = .Fields("ssShippingMethodID").Value
  %>
  <tr class="Inactive" onmousedown="highlightRow(this,true);">
    <td align="center"><input type="checkbox" NAME="dirty" ID="dirty" value="<%= plngShippingMethodID  %>"><input type="hidden" NAME="ShippingMethodID" ID="ShippingMethodID" value="<%= plngShippingMethodID  %>"></td>
    <td align="right"><%= plngShippingMethodID %>&nbsp;</td>
    <td align="center"><input type="text" name="ssShippingMethodName<%= plngShippingMethodID %>" id="Text2" value="<%= Trim(.Fields("ssShippingMethodName").Value) %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isEmpty(this, false, 'Please enter a Shipping Method Name')" size="20"></td>
    <td align="center"><input type="text" name="ssShippingMethodCode<%= plngShippingMethodID %>" id="Text2" value="<%= Trim(.Fields("ssShippingMethodCode").Value) %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isEmpty(this, false, 'Please enter a Shipping Method Code')" size="20"></td>
    <td align="center"><select name="ssShippingCarrierID<%= plngShippingMethodID %>" id="ssShippingCarrierID<%= plngShippingMethodID %>"  onchange="makeDirty(this, <%= plngShippingMethodID %>);"><% Call MakeCombo(mobjrsCarriers,"ssShippingCarrierName","ssShippingCarrierID",.Fields("ssShippingCarrierID").Value) %></select></td>
    <td align="center"><input type="checkbox" NAME="ssShippingMethodEnabled<%= plngShippingMethodID %>" ID="Checkbox1" value="1" <% WriteCheckboxValue(.Fields("ssShippingMethodEnabled").Value) %> onclick="makeDirty(this, <%= plngShippingMethodID %>);"></td>
    <td align="center"><input type="checkbox" NAME="ssShippingMethodIsSpecial<%= plngShippingMethodID %>" ID="Checkbox2" value="1" <% WriteCheckboxValue(.Fields("ssShippingMethodIsSpecial").Value) %> onclick="makeDirty(this, <%= plngShippingMethodID %>);"></td>
    <td align="center"><input type="text" name="ssShippingMethodMinCharge<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMinCharge").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="4" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodMultiple<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMultiple").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="4" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodPerPackageFee<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodPerPackageFee").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="4" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodPerShipmentFee<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodPerShipmentFee").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="4" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodOfferFreeShippingAbove<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodOfferFreeShippingAbove").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, true, 'Please enter a number')" size="4" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodLimitFreeShippingByWeight<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodLimitFreeShippingByWeight").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="6" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodMinWeight<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMinWeight").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="6" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodMaxWeight<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMaxWeight").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="6" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodPrefWeight<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodPrefWeight").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="6" style="text-align: right"></td>
    <td align="center"><input type="text" name="ssShippingMethodMaxGirth<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMaxGirth").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="6" style="text-align: right"></td>
    <td align="center"><input type="radio" name="ssShippingMethodDefault" id="ssShippingMethodDefault<%= plngShippingMethodID %>" value="<%= plngShippingMethodID %>" <% WriteCheckboxValue(.Fields("ssShippingMethodDefault").Value) %> onchange="makeDirty(this, <%= plngShippingMethodID %>);"></td>
    <td align="center">
		<select name="ssShippingMethodCountryRule<%= plngShippingMethodID %>" id="ssShippingMethodCountryRule<%= plngShippingMethodID %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);">
		  <option value="0" <%= isSelected(.Fields("ssShippingMethodCountryRule").Value=0) %>></option>
		  <option value="1" <%= isSelected(.Fields("ssShippingMethodCountryRule").Value=1) %>>U.S.</option>
		  <option value="2" <%= isSelected(.Fields("ssShippingMethodCountryRule").Value=2) %>>Int'l</option>
		  <option value="3" <%= isSelected(.Fields("ssShippingMethodCountryRule").Value=3) %>>All</option>
		</select>
    </td>
    <!--
    <td align="center"><input type="text" name="ssShippingMethodMaxLength<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMaxLength").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="20"></td>
    <td align="center"><input type="text" name="ssShippingMethodMaxWidth<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMaxWidth").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="20"></td>
    <td align="center"><input type="text" name="ssShippingMethodMaxHeight<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodMaxHeight").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="20"></td>
    <td align="center"><input type="text" name="ssShippingMethodClass<%= plngShippingMethodID %>" id="Text2" value="<%= .Fields("ssShippingMethodClass").Value %>" onchange="makeDirty(this, <%= plngShippingMethodID %>);" onblur="return isNumeric(this, false, 'Please enter a number')" size="20"></td>
    <td align="center"><input type="checkbox" name="ssShippingMethodLocked<%= plngShippingMethodID %>" ID="ssShippingMethodLocked<%= plngShippingMethodID %>" value="1" <% WriteCheckboxValue(.Fields("ssShippingMethodLocked").Value) %>></td>
	-->
  </tr>

  <%
	  .MoveNext
	Loop
  End With
  %>
 <tr class="tblhdr">
	<th align="left" colspan="19">
	  &nbsp;&nbsp;<input class="butn" id="btnCheckAll" name="btnCheckAll" type="button" value="Check All" onclick="CheckAll(true);">
	  &nbsp;&nbsp;<input class="butn" id="btnUnCheckAll" name="btnUnCheckAll" type="button" value="Uncheck All" onclick="CheckAll(false);">
	  <input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/PostageRate/help_PostageRate.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
      <input class="butn" title="Add a new carrier" id="btnAdd" name="btnAdd" type="button" onclick="this.form.Action.value='Add'; this.form.submit();" value="Add New">
      <input class="butn" title="delete checked carriers" id="btnDelete" name="btnDelete" type="button" onclick="this.form.Action.value='Delete'; this.form.submit();" value="Delete">
      <input class="butn" title="move selected carrier up" id="btnMoveUp" name="btnMoveUp" type="button" onclick="moveCarrier(true);" value="Move Up" disabled>
      <input class="butn" title="move selected carrier down" id="btnMoveDown" name="btnMoveDown" type="button" onclick="moveCarrier(false);" value="Move Down" disabled>
      <input class="butn" title="Save changes" id="btnUpdate" name="btnUpdate" type="submit" value="Save">
    </th>
  </TR>
  <tr>
    <td colspan="18" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" ID="Table1">
        <tr>
          <td align="center"><a href="ssPostageRate_ShippingCarriersAdmin.asp">Configure Shipping Carriers</a></td>
        </tr>
      </table>
  
    </td>
  </tr>
</TABLE>

</FORM>

</CENTER>
</BODY>
</HTML>
<%
    Set mclsShipMethod = Nothing

	Call ReleaseObject(mobjrsCarriers)
	Call ReleaseObject(cnn)

    Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Function WriteCheckboxValue(vntValue)

	If len(Trim(vntValue) & "") > 0 Then
		If cBool(vntValue) Then Response.Write "CHECKED"
	End If

End Function	'WriteCheckboxValue

'************************************************************************************************************************************

%>