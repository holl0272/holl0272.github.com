<%Option Explicit
'********************************************************************************
'*   Product Placement Tool For StoreFront 5.0			                        *
'*   Release Version:	1.00.001		                                        *
'*   Release Date:		March 1, 2004											*
'*   Revision Date:		March 1, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.00.001 (March 1, 2004)			                                *
'*     - Initial Release												        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Dim mlngPageCount
Dim mlngAbsolutePage
Dim mlngMaxRecords

mlngMaxRecords = LoadRequestValue("PageSize")

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Server.ScriptTimeout = 300			'in seconds. Adjust for large databases or if some products have a lot of attributes. Server Default is usually 90 seconds
If len(mlngMaxRecords) = 0 Then mlngMaxRecords = 50	'Set your default Maximum Records to show in summary table

'/
'/////////////////////////////////////////////////

Class clsProduct
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsProducts
Private pblnError
Private cblnDebug
Private plngProductPlacementType

'database variables

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    cblnDebug = False
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

Public Property Let ProductPlacementType(byVal lngType)
	plngProductPlacementType = lngType
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

Public Function AECategoryFilter(ByVal bytCategoryFilter, ByVal bytsubCategoryFilter)

Dim sql
Dim prsAEProductCategories
Dim i

	sql = "SELECT sfCategories.catID, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.bottom" _
		& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
		& " ORDER BY sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName"
	'debugprint "sql",sql
	
	Set prsAEProductCategories = GetRS(sql)
	With prsAEProductCategories
		If Not .EOF Then

			'this creates the first dropdown
			Dim pstrTemp
			Dim pstrCatValue
			Dim pstrCatName
			Dim pbytCatDepth
			Dim pstrsubCatValue
			Dim pstrsubCatName
			Dim pstrSelected

			If Len(bytCategoryFilter) = 0 And Len(bytsubCategoryFilter) = 0 Then pstrSelected = "selected"
			pstrTemp = "<select id='CategoryFilter' name='CategoryFilter' size='10' multiple>" & vbcrlf _
					 & "  <option value='.' " & pstrSelected & ">- All -</Option>" & vbcrlf

			For i=1 to .RecordCount

				If len(.Fields("Depth").value & "") > 0 Then 'Then No sub-categories so don't display
					pbytCatDepth = .Fields("Depth").value
					
					If pstrCatName <> Trim(.Fields("catName").value) Then
						pstrCatValue = .Fields("catID").value & "."
						pstrCatName = .Fields("catName").value
						pstrSelected = ""
						If pstrCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
						pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrCatValue & chr(34) & pstrSelected & ">" & pstrCatName & "</option>" & vbcrlf 
						
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					Else
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					End If
				End If
				
				.MoveNext
			Next
			pstrTemp = pstrTemp & "</select>" & vbcrlf
			
			
		Else
			pstrTemp = "<font color=red>No Categories</font><br />"
			pstrTemp = pstrTemp & "<select id=CatSource name=CatSource size=10 multiple>"
			pstrTemp = pstrTemp & "</select>"
		End If
	End With

	AECategoryFilter = pstrTemp
	
End Function	'AECategoryFilter

'***********************************************************************************************

Public Function Load()

dim pstrSQL
dim p_strWhere
dim i
dim sql
Dim plngNumRecords

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		If cblnSF5AE Then
			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, subcatDetailID, sfSubCatDetail.sortCatDetail, sfProducts.sortCat, sfProducts.sortMfg, sfProducts.sortVend" _
					& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
					& " " & mstrSQLWhere _
					& " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, subcatDetailID, sfSubCatDetail.sortCatDetail, sfProducts.sortCat, sfProducts.sortMfg, sfProducts.sortVend" _
					& mstrsqlHaving _
					& " ORDER BY sfProducts.prodID"
		Else
			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, sfProducts.sortCat, sfProducts.sortMfg, sfProducts.sortVend" _
					& " FROM sfProducts " _
					& " " & mstrSQLWhere _
					& " ORDER BY sfProducts.prodID"

		End If
		pstrSQL = adjustProductSort(pstrSQL, plngProductPlacementType)

		If cblnDebug Then Response.Flush
		  
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		
		On Error Resume Next
		'.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		If cblnDebug Then debugprint "Used pstrSQL",pstrSQL
		.Open pstrSQL, cnn, 3,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			Response.Write "<div class='FatalError'>You may need to upgrade your database to use the Product Pricing Tool</div>" _
							& "<h3><a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=ProductPlacement'>Click here to upgrade</a></h3>"
			Response.Write "<font color=red>Error in Load: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			Response.Write "<font color=red>Error in Load: sql = " & pstrSQL & "</font><br />" & vbcrlf
			Call ShowStoreFrontVersion
			Response.Flush
			Err.Clear
			Load = False
			Exit Function
		End If
		
		If cblnDebug Then
			debugprint ".EOF",.EOF
			debugprint ".RecordCount",.RecordCount
		End If
		If .EOF Then
			mlngPageCount = 0
		Else
			mlngPageCount = .PageCount
		End If
		If cblnDebug Then debugprint "mlngPageCount", mlngPageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
	end with

    Load = (Not prsProducts.EOF)

End Function    'Load

'***********************************************************************************************

Public Function Update()

Dim pstrSQL
Dim strErrorMessage
Dim vItem
Dim paryProductIDs
Dim paryProductPlacementPos
Dim i
Dim pstrProdID
Dim plngValue

'On Error Resume Next

    pblnError = False

	paryProductIDs = Split(Request.Form("fieldID"),",")
	paryProductPlacementPos = Split(Request.Form("ProductPlacementPos"),",")
	
	'Update Product Prices
	'For Each vItem In Request.Form
	'	Response.Write vItem & " = " & Request.Form(vItem) & "<br />"
	'Next
	'Response.Flush
	'Update the products
	For i = 0 To UBound(paryProductIDs)
		pstrProdID = Trim(paryProductIDs(i))
		plngValue = Trim(paryProductPlacementPos(i))
		If Len(plngValue) = 0 Then plngValue = i
		If plngProductPlacementType = 3 Then
			pstrSQL = "Update sfSubCatDetail Set " & ProductSortField(plngProductPlacementType) & "=" & plngValue & " Where subcatDetailID=" & pstrProdID
		Else
			pstrSQL = "Update sfProducts Set " & ProductSortField(plngProductPlacementType) & "=" & plngValue & " Where prodID='" & Replace(pstrProdID, "'", "''") & "'"
		End If
		'Response.Write i & ": " & pstrSQL & "<br />"
		cnn.Execute pstrSQL,,128
	Next 'i

    Update = (not pblnError)

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

End Class   'clsProduct

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="Common/ssProduct_CommonFilter.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
cblnSF5AE = False
'******************************************************************************************************************************************************************

mstrPageTitle = "Product Placement Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsProduct
Dim mlngProductPlacementType
Dim mblnShowFilter, mblnShowSummary, mstrShow


	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	mAction = LoadRequestValue("Action")
	If Len(mAction) = 0 Then mAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
    Call LoadFilter

	'Necessary since the ALL category filter can come in as ...
	If mlngCategoryFilter = "..." Then
		mlngProductPlacementType = DetermineSortType("", mlngManufacturerFilter, mlngVendorFilter)
	Else
		mlngProductPlacementType = DetermineSortType(mlngCategoryFilter, mlngManufacturerFilter, mlngVendorFilter)
	End If    
	
    Set mclsProduct = New clsProduct
    With mclsProduct
		If mAction = "Update" Then
			.ProductPlacementType = Request.Form("ProductPlacementType")
			.Update
		End If
		.ProductPlacementType = mlngProductPlacementType
		.Load
	End With
    
	Call WriteHeader("",True)
%>
<script LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></script>
<script LANGUAGE=javascript>
<!--

function CheckAll(blnCheck)
{
	var plngCount;
	var i;

	plngCount = document.frmData.prodID.length;
	if (document.frmData.prodID.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
		document.frmData.prodID[i].checked = blnCheck;
		}
	}else{
	document.frmData.prodID.checked = blnCheck;
	}
	
	plngCount = document.frmData.attrdtID.length;
	if (document.frmData.attrdtID.checked==undefined)
	{
		for (i=0; i < plngCount;i++)
		{
		document.frmData.attrdtID[i].checked = blnCheck;
		}
	}else{
	document.frmData.attrdtID.checked = blnCheck;
	}
	
}

function moveRow(blnUp, rowPos)
{
var tblSummary = document.all("tblSummary");
var plngCount = tblSummary.rows.length-2;
var i;
	
	if (blnUp)
	{
		tblSummary.rows[mlngSelectedRow].swapNode(tblSummary.rows[mlngSelectedRow-1]);
		if (mlngSelectedRow==plngCount)
		{
			document.frmData.btnMoveDown[plngCount-1].disabled = true;
			document.frmData.btnMoveDown[plngCount-2].disabled = false;
		}
		mlngSelectedRow = mlngSelectedRow - 1;
		if (mlngSelectedRow==1)
		{
			document.frmData.btnMoveUp[0].disabled = true;
			document.frmData.btnMoveUp[1].disabled = false;
		}
	}else{
		tblSummary.rows[mlngSelectedRow].swapNode(tblSummary.rows[mlngSelectedRow+1]);
		if (mlngSelectedRow==1)
		{
			document.frmData.btnMoveUp[0].disabled = true;
			document.frmData.btnMoveUp[1].disabled = false;
		}
		mlngSelectedRow = mlngSelectedRow + 1;
		if (mlngSelectedRow==plngCount)
		{
			document.frmData.btnMoveDown[plngCount-1].disabled = true;
			document.frmData.btnMoveDown[plngCount-2].disabled = false;
		}
	}
}

var mlngSelectedRow;

function setRow(theRow, blnHighlight)
{
var tblSummary = document.all("tblSummary");
var plngCount = tblSummary.rows.length - 3;
var i;

	mlngSelectedRow = theRow.rowIndex;

	for (i=1; i < tblSummary.rows.length-1;i++)
	{
		tblSummary.rows[i].className = "Inactive";
	}
	
	if (blnHighlight)
	{
		theRow.className = "Selected";
	}else{
		theRow.className = "Inactive";
	}
}

function setOrder()
{
var i;

	for (i=0; i < document.frmData.ProductPlacementPos.length;i++)
	{
		document.frmData.ProductPlacementPos[i].value = i+1;
	}
}
//-->
</script>

<center>

<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
    <th>&nbsp;</th>
    <th align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br />
	</th>
  </tr>
</table>

<form action="ssProductPlacementAdmin.asp" id="frmData" name="frmData" method="post">
<input type="hidden" id="Action" name="Action" value="Update">
<input type="hidden" id="blnShowFilter" name="blnShowFilter" value="Update">
<input type="hidden" id="ProductPlacementType" name="ProductPlacementType" value="<%= mlngProductPlacementType %>">


<% Call WriteProductFilter(False) %>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblSummary">
  <tr>
    <th align="left">Product ID</th>
    <th align="left">Product</th>
    <th>Price</th>
    <th>Sort by <%= ProductSortName(mlngProductPlacementType) %></th>
    <th>&nbsp;</th>
  </tr>
  <% 
  Dim pstrPrevID
  Dim pstrFieldID
  Dim plngUniqueProducts
  Dim pstrPrice
  Dim lngRecordCounter
  Dim mstrMoveUpDisabled
  Dim mstrMoveDownDisabled
  
 
  plngUniqueProducts = 0
  
  With mclsProduct.rsProducts 
	For lngRecordCounter = 1 To .RecordCount
		If .Fields("prodID").Value <> pstrPrevID Then
			pstrPrevID = .Fields("prodID").Value
			plngUniqueProducts = plngUniqueProducts + 1
			mstrMoveUpDisabled = ""
			mstrMoveDownDisabled = ""
			If lngRecordCounter = 1 Then mstrMoveUpDisabled = "disabled"
			If lngRecordCounter = .RecordCount Then mstrMoveDownDisabled = "disabled"
		  
		  If .Fields("prodSaleIsActive").Value = "1" Then
			pstrPrice = "<font color=red><strike>" & WriteCurrency(.Fields("prodPrice").Value) & "</strike> " & WriteCurrency(.Fields("prodSalePrice").Value) & "</font>"
		  Else
			pstrPrice = WriteCurrency(.Fields("prodPrice").Value)
		  End If
		  
		If mlngProductPlacementType = 3 Then
			pstrFieldID = .Fields("subcatDetailID").Value
		Else
			pstrFieldID = pstrPrevID
		End If
			
'  <tr><td colspan="6"><hr width="95%"</td></tr>
  %>
  <tr id="row<%= lngRecordCounter %>" onmousedown="setRow(this,true);">
    <td><%= pstrPrevID %>&nbsp;</td>
    <td><%= .Fields("prodName").Value %>&nbsp;</td>
    <td align="center"><%= pstrPrice %>&nbsp;</td>
    <td align="center">
	<input type="hidden" id="fieldID" name="fieldID" value="<%= pstrFieldID %>">
    <input style="text-align: center;" type="text" name="ProductPlacementPos" id="ProductPlacementPos" value="<%= .Fields(ProductSortField(mlngProductPlacementType)).Value %>" onblur="return isNumeric(this, true, 'Please enter a number')" size="5"></td>
    <td>
      <input class="butn" title="move up" id="btnMoveUp" name="btnMoveUp" type="button" onclick="moveRow(true);" value="Move Up" <%= mstrMoveUpDisabled %>>
      <input class="butn" title="move down" id="btnMoveDown" name="btnMoveDown" type="button" onclick="moveRow(false);" value="Move Down" <%= mstrMoveDownDisabled %>>
    </td>
  </tr>
  <% 
		End If
	  .MoveNext
	Next	'lngRecordCounter
  End With
  %>
 <tr class="tblhdr">
	<th align="center" colspan="5">Products Returned: <%= plngUniqueProducts %></th>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
		<input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/ProductPricingTool_help.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
        <input class="butn" title="set order" id="btnSetOrder" name="btnSetOrder" type="button" value="Set Order" onclick="setOrder(); alert('The changes will not be saved until you press the update button.');">
        <input class="butn" title="Save changes" id="btnUpdate" name="btnUpdate" type="submit" value="Save Changes">
    </td>
  </tr>
</table>

</form>

</center>
</BODY>
</HTML>
<%

	Call ReleaseObject(cnn)

    Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Function DetermineSortType(byVal lngCategoryFilter, byVal lngManufacturerFilter, byVal lngVendorFilter)
'Options include
'0 - Category
'1 - Manufacturer
'2 - Vendor
'3 - Category from sfSubCatDetail

Dim plngTemp

	If Len(lngCategoryFilter) > 0 Then
		If cblnSF5AE Then
			plngTemp = 3
		Else
			plngTemp = 0
		End If
	ElseIf Len(lngManufacturerFilter) > 0 Then
		plngTemp = 1
	ElseIf Len(lngVendorFilter) > 0 Then
		plngTemp = 2
	Else
		plngTemp = 0
	End If

	DetermineSortType = plngTemp
	
End Function	'DetermineSortType

'************************************************************************************************************************************

Function adjustProductSort(byVal strSQL, byVal sortTypePrimary)

Dim pstrSQL
Dim pstrSQL_OrderByIncoming
Dim pstrSQL_OrderBy
Dim plngPos

	plngPos = InStr(LCase(strSQL), "order by")
	
	If plngPos > 0 Then
		pstrSQL = Left(strSQL, plngPos - 1)
		pstrSQL_OrderByIncoming = Right(strSQL, Len(strSQL) - plngPos - 8)
	Else
		pstrSQL = strSQL
	End If

	pstrSQL_OrderBy = ProductSortField(sortTypePrimary)
	
	If Len(pstrSQL_OrderByIncoming) > 0 Then
		pstrSQL_OrderBy = "Order By " & pstrSQL_OrderBy & "," & pstrSQL_OrderByIncoming
	Else
		pstrSQL_OrderBy = "Order By " & pstrSQL_OrderBy
	End If

	pstrSQL = pstrSQL & pstrSQL_OrderBy
	
	adjustProductSort = pstrSQL

End Function	'adjustProductSort

'************************************************************************************************************************************

Function ProductSortField(byVal sortTypePrimary)

	Select Case sortTypePrimary
		Case 0: ProductSortField = "sortCat"
		Case 1: ProductSortField = "sortMfg"
		Case 2: ProductSortField = "sortVend"
		Case 3: ProductSortField = "sortCatDetail"
		Case Else: ProductSortField = "sortCat"
	End Select
	
End Function	'ProductSortField

'************************************************************************************************************************************

Function ProductSortName(byVal sortTypePrimary)

	Select Case sortTypePrimary
		Case 0: ProductSortName = "Category"
		Case 1: ProductSortName = "Manufacturer"
		Case 2: ProductSortName = "Vendor"
		Case 3: ProductSortName = "Category - AE"
		Case Else: ProductSortName = "Unknown"
	End Select
	
End Function	'ProductSortName

'************************************************************************************************************************************

%>
