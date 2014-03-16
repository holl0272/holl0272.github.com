<%Option Explicit
'********************************************************************************
'*   Product Pricing Tool For StoreFront 5.0			                        *
'*   Release Version:	1.00.003		                                        *
'*   Release Date:		October 3, 2002											*
'*   Revision Date:		May 23, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.00.003 (May 23, 2004)			                                *
'*     - Added feature to require filter selection before displaying records    *
'*                                                                              *
'*   Version 1.02 (July 18, 2003)				                                *
'*     - Added feature to show records when attribute join doesn't work         *
'*       Still don't know why sometimes this fails						        *
'*                                                                              *
'*   Version 1.01 (March 15, 2003)				                                *
'*     - Added feature to disable attributes                                    *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
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

Private pblnUseAttributes

'database variables

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    pblnUseAttributes = True
    'pblnUseAttributes = False
    cblnDebug = False
    'cblnDebug = True
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsProducts)
End Sub

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

Public Property Get UseAttributes()
    UseAttributes = pblnUseAttributes
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
dim pstrSQLAttr
Dim plngNumRecords

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		If cblnSF5AE Then
			pstrSQLAttr = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, sfAttributes.attrName, sfAttributeDetail.attrdtID, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType" _
						& " FROM (((sfProducts LEFT JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId) INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
						& " " & mstrSQLWhere _
						& " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, sfAttributes.attrName, sfAttributeDetail.attrdtID, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType, sfAttributeDetail.attrdtOrder" _
						& mstrsqlHaving _
						& " ORDER BY sfProducts.prodID, sfAttributes.attrName, sfAttributeDetail.attrdtOrder"

			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive" _
					& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
					& " " & mstrSQLWhere _
					& " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive" _
					& mstrsqlHaving _
					& " ORDER BY sfProducts.prodID"

		Else
			pstrSQLAttr = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive, sfAttributes.attrName, sfAttributeDetail.attrdtID, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType" _
						& " FROM (sfProducts LEFT JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
						& " " & mstrSQLWhere _
						& " ORDER BY sfProducts.prodID, sfAttributes.attrName, sfAttributeDetail.attrdtOrder"

			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodSaleIsActive" _
					& " FROM sfProducts " _
					& " " & mstrSQLWhere _
					& " ORDER BY sfProducts.prodID"

		End If
		
		pstrSQL = Replace(pstrSQL, "sfProducts.prodPrice, sfProducts.prodSalePrice,", "sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice,")
		pstrSQLAttr = Replace(pstrSQLAttr, "sfProducts.prodPrice, sfProducts.prodSalePrice,", "sfProducts.prodPrice, sfProducts.prodSalePrice, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice,")

		'If cblnDebug Then Response.Flush
		  
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		
		'On Error Resume Next
		'.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		
		If pblnUseAttributes Then
			.Open pstrSQLAttr, cnn, 3,1	'adOpenKeySet,adLockReadOnly
			
			If .EOF Then
				.Close
				pblnUseAttributes = False
				If cblnDebug Then debugprint "Used pstrSQL",pstrSQL
				.Open pstrSQL, cnn, 3,1	'adOpenKeySet,adLockReadOnly
			Else
				On Error Resume Next
				
				plngNumRecords = .RecordCount
				
				If Err.number <> 0 Then
					Err.Clear
					.Close
					pblnUseAttributes = False
					If cblnDebug Then debugprint "Used pstrSQL",pstrSQL
					.Open pstrSQL, cnn, 3,1	'adOpenKeySet,adLockReadOnly
				Else
					If cblnDebug Then debugprint "Used pstrSQLAttr",pstrSQLAttr
				End If
				
			End If
			
		Else
			If cblnDebug Then debugprint "Used pstrSQL",pstrSQL
			.Open pstrSQL, cnn, 3,1	'adOpenKeySet,adLockReadOnly
		End If
		
		If Err.number <> 0 Then
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

Dim sql
Dim strErrorMessage
Dim vItem
Dim paryProductIDs
Dim paryAttributeIDs
Dim i
Dim pstrProdID
Dim pstrProdPrice
Dim pstrProdSalePrice
Dim pstrProdPLPrice
Dim pstrSaleIsActive

'On Error Resume Next

    pblnError = False

	paryProductIDs = Split(Request.Form("prodID"),",")
	paryAttributeIDs = Split(Request.Form("attrdtID"),",")
	
	'Update the products
	For i = 0 To UBound(paryProductIDs)
		pstrProdID = Trim(paryProductIDs(i))
		pstrProdPrice = Trim(Request.Form("prodPrice" & pstrProdID))
		pstrProdSalePrice = Trim(Request.Form("prodSalePrice" & pstrProdID))
		pstrProdPLPrice = Trim(Request.Form("prodPLPrice" & pstrProdID))
		pstrProdPLPrice = Replace(pstrProdPLPrice, " ", "")
		pstrProdPLPrice = Replace(pstrProdPLPrice, ",", ";")
		
		If Len(Trim(Request.Form("prodSaleIsActive" & pstrProdID))) > 0 Then
			pstrSaleIsActive = "1"
		Else
			pstrSaleIsActive = "0"
		End If

		If Len(pstrProdSalePrice) > 0 Then
			If cblnAddon_ProductPricing Then
				sql = "Update sfProducts Set prodPrice='" & pstrProdPrice & "', prodPLPrice='" & pstrProdPLPrice & "', prodSaleIsActive='" & pstrSaleIsActive & "', prodSalePrice='" & pstrProdSalePrice & "' Where prodID='" & pstrProdID & "'"
			Else
				sql = "Update sfProducts Set prodPrice='" & pstrProdPrice & "', prodSaleIsActive='" & pstrSaleIsActive & "', prodSalePrice='" & pstrProdSalePrice & "' Where prodID='" & pstrProdID & "'"
			End If
		Else
			If cblnAddon_ProductPricing Then
				sql = "Update sfProducts Set prodPrice='" & pstrProdPrice & "', prodPLPrice='" & pstrProdPLPrice & "', prodSaleIsActive='" & pstrSaleIsActive & "' Where prodID='" & pstrProdID & "'"
			Else
				sql = "Update sfProducts Set prodPrice='" & pstrProdPrice & "', prodSaleIsActive='" & pstrSaleIsActive & "' Where prodID='" & pstrProdID & "'"
			End If
		End If
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
	Next 'i
	
	Dim plngAttrID
	Dim pstrAttrPrice
	Dim pbytAttrType
	
	'Update the attributes
	For i = 0 To UBound(paryAttributeIDs)
		plngAttrID = Trim(paryAttributeIDs(i))
		pstrAttrPrice = Trim(Request.Form("attrdtPrice" & plngAttrID))
		pbytAttrType = Trim(Request.Form("attrdtType" & plngAttrID))
		
		sql = "Update sfAttributeDetail Set attrdtPrice='" & pstrAttrPrice & "', attrdtType=" & pbytAttrType & " Where attrdtID=" & plngAttrID
		'Response.Write i & ": " & sql & "<br />"
		cnn.Execute sql,,128
	Next 'i
	
	'Update Product Prices
	For Each vItem In Request.Form
	'	Response.Write vItem & " = " & Request.Form(vItem) & "<br />"
	Next


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
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'******************************************************************************************************************************************************************

mstrPageTitle = "Product Pricing Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Const cblnRequireInitialFilter = True
Dim mAction
Dim mclsProduct
Dim i

Dim mblnShowFilter, mblnShowSummary, mstrShow

	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	
	mAction = LoadRequestValue("Action")
	mblnShowSummary = CBool(Len(mAction) > 0) Or (CBool(Len(mAction) = 0) And Not cblnRequireInitialFilter)
	If Len(mAction) = 0 Then mAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	
    Call LoadFilter
    
	dim mrsCategory, mrsVendor, mrsManufacturer
	
	Set mrsCategory = GetRS("Select catID,catName from sfCategories Order By catName")
	Set mrsVendor = GetRS("Select vendID,vendName from sfVendors Order By vendName")
	Set mrsManufacturer = GetRS("Select mfgID,mfgName from sfManufacturers Order By mfgName")

    Set mclsProduct = New clsProduct
    With mclsProduct
		If mAction = "Update" Then .Update
		If mblnShowSummary Then mblnShowSummary = .Load
	End With
   
	Call WriteHeader("",True)
%>
<script language=javascript>
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

function checkAll_custom(strType, blnCheck)
{
	var plngCount;
	var i;
	var strCheckboxName;

	switch (strType)
	{
		case 'update':
			checkAll(document.frmData.prodID, blnCheck)
			checkAll(document.frmData.attrdtID, blnCheck)
			break;
		case 'prodSaleIsActive':
			plngCount = document.frmData.prodID.length;
			if (document.frmData.prodID.checked==undefined)
			{
				for (i=0; i < plngCount;i++)
				{
				strCheckboxName = strType + document.frmData.prodID[i].value;
				checkAll(document.all(strCheckboxName), blnCheck)
				}
			}else{
				strCheckboxName = strType + document.frmData.prodID.value;
				checkAll(document.all(strCheckboxName), blnCheck)
			}
			break;
		case 'IsActive':
			plngCount = document.frmData.prodID.length;
			if (document.frmData.prodID.checked==undefined)
			{
				for (i=0; i < plngCount;i++)
				{
				strCheckboxName = strType + document.frmData.prodID[i].value;
				checkAll(document.all(strCheckboxName), blnCheck)
				}
			}else{
				strCheckboxName = strType + document.frmData.prodID.value;
				checkAll(document.all(strCheckboxName), blnCheck)
			}
			break;
	}
	return true;

	if (document.frmData.attrdtID != undefined)
	{
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
	
}

//-->
</script>

<center>

<table border=0 cellpadding=5 cellspacing=1 width="95%" id="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
    <th>&nbsp;</th>
    <th align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onmouseover="return DisplayTitle(this);" onmouseout="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br />
	</th>
  </tr>
</table>

<form action="ssProductPricingTool.asp" id="frmData" name="frmData" method="post">
<input type="hidden" id="Action" name="Action" value="Update">
<input type="hidden" id="blnShowFilter" name="blnShowFilter" value="Update">

<% Call WriteProductFilter(False) %>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblSummary">
  <tr class="hdrNonSelected">
    <th>Update<input type="checkbox" name="chkCheckAll_Update" id="chkCheckAll_Update1"  onclick="checkAll_custom('update', this.checked); checkAll(this.form.chkCheckAll_Update2, this.checked);" value="" title="Toggle update checkboxes"></th>
    <th align="left">Product ID</th>
    <th align="left">Product</th>
    <th>Regular Price</th>
    <th>Sale Price</th>
    <%
		If hasProductPricingLevels Then
			For i = 0 To UBound(maryPricingLevels)
				Response.Write "<th>" & maryPricingLevels(i) & "</th>"
			Next 'i
		End If	'hasProductPricingLevels
    %>
    <th style="cursor:auto;">On Sale&nbsp;<input type="checkbox" name="chkCheckAll_prodSaleIsActive" id="chkCheckAll_prodSaleIsActive1"  onclick="checkAll_custom('prodSaleIsActive', this.checked); checkAll(this.form.chkCheckAll_prodSaleIsActive2, this.checked);" value="" title="Toggle on sale checkboxes"></th>
  </tr>
  <% 
  Dim pstrPrevID
  Dim pstrPrevAttrID
  Dim plngUniqueProducts
  plngUniqueProducts = 0
  
  If mblnShowSummary Then
  With mclsProduct.rsProducts 
	Do While Not .EOF
		If .Fields("prodID").Value <> pstrPrevID Then
		  pstrPrevID = .Fields("prodID").Value
		  plngUniqueProducts = plngUniqueProducts + 1
'  <tr><td colspan="6"><hr width="95%"</td></tr>
  %>
  <tr>
    <td align="center"><input type="checkbox" name="prodID" id="prodID" value="<%= pstrPrevID %>"></td>
    <td><%= pstrPrevID %>&nbsp;</td>
    <td><%= .Fields("prodName").Value %>&nbsp;</td>
    <td align="center"><input type="text" class="priceBox" name="prodPrice<%= pstrPrevID %>" id="prodPrice<%= pstrPrevID %>" value="<%= .Fields("prodPrice").Value %>" onchange="makeDirty('<%= pstrPrevID %>', false);" onblur="return isNumeric(this, false, 'Please enter a number')" size="8"></td>
    <td align="center"><input type="text" class="priceBox" name="prodSalePrice<%= pstrPrevID %>" id="prodSalePrice<%= pstrPrevID %>" value="<%= .Fields("prodSalePrice").Value %>" onchange="makeDirty('<%= pstrPrevID %>', false);" onblur="return isNumeric(this, false, 'Please enter a number')" size="8"></td>
    <%
    Dim maryPLPrices
    Dim mstrPLPrice
		If hasProductPricingLevels Then
			maryPLPrices = Split(.Fields("prodPLPrice").Value & "",";")
			For i = 0 To UBound(maryPricingLevels)
				If i > UBound(maryPLPrices) Then
					mstrPLPrice = ""
				Else
					mstrPLPrice = Trim(maryPLPrices(i))
				End If
				%><td align="center"><input type="text" class="priceBox" name="prodPLPrice<%= pstrPrevID %>" id="prodPLPrice<%= pstrPrevID & i %>" value="<%= mstrPLPrice %>" onchange="makeDirty('<%= pstrPrevID %>', false);" onblur="return isNumeric(this, true, 'Please enter a number')" size="8"></td><%
			Next 'i
		End If	'hasProductPricingLevels
    %>
    <td align="center"><input type="checkbox" name="prodSaleIsActive<%= pstrPrevID %>" id="prodSaleIsActive<%= pstrPrevID %>" value="1" <% WriteCheckboxValue(.Fields("prodSaleIsActive").Value) %> onclick="makeDirty('<%= pstrPrevID %>', false);"></td>
  </tr>
  <% 
		End If
		If mclsProduct.UseAttributes Then
		If .Fields("attrdtID").Value <> pstrPrevAttrID Then
			pstrPrevAttrID = .Fields("attrdtID").Value
  %>
  <tr>
    <td align="center"><input type="checkbox" name="attrdtID" id="attrdtID" value="<%= .Fields("attrdtID").Value %>"></td>
    <td>&nbsp;</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;<%= .Fields("attrName").Value %>:&nbsp;<%= .Fields("attrdtName").Value %></td>
    <td align="center">
    <input type="text" class="priceBox" name="attrdtPrice<%= .Fields("attrdtID").Value %>" id="attrdtPrice<%= .Fields("attrdtID").Value %>" value="<%= .Fields("attrdtPrice").Value %>" onchange="makeDirty('<%= .Fields("attrdtID").Value %>', true);" onblur="return isNumeric(this, false, 'Please enter a number')" size="20" tag="<%= pstrPrevID %>"></td>
    <td>
		<select size="1" name="attrdtType<%= .Fields("attrdtID").Value %>" id="attrdtType<%= .Fields("attrdtID").Value %>" onchange="makeDirty('<%= .Fields("attrdtID").Value %>', true);">
			<option <% If Trim(.Fields("attrdtType").Value & "") = "1" Then Response.Write "selected" %> value="1">Increase</option>
			<option <% If Trim(.Fields("attrdtType").Value & "") = "2" Then Response.Write "selected" %> value="2">Decrease</option>
			<option <% If Trim(.Fields("attrdtType").Value & "") = "0" Then Response.Write "selected" %> value="0">No Change</option>
		</select>
	</td>
    <td>&nbsp;</td>
  </tr>
  <%
      End If
	  End If	'.UseAttributes
	  .MoveNext
	Loop
  End With
  Else
  %>
  <tr><td colspan="6"><hr /><h3>No products meet the filter criteria or you have not selected a filter criteria</h3><hr /></td></tr>
  <%
  End If	'mblnShowSummary
  %>
 <tr class="tblhdr">
	<th align="center" colspan="1"><input type="checkbox" name="chkCheckAll_Update" id="chkCheckAll_Update2"  onclick="checkAll_custom('update', this.checked); checkAll(this.form.chkCheckAll_Update1, this.checked);" value="" title="Toggle update checkboxes"></th>
	<th align="left" colspan="4">Products Returned: <%= plngUniqueProducts %></th>
	<th align="center" colspan="1"><input type="checkbox" name="chkCheckAll_prodSaleIsActive" id="chkCheckAll_prodSaleIsActive2"  onclick="checkAll_custom('prodSaleIsActive', this.checked); checkAll(this.form.chkCheckAll_prodSaleIsActive1, this.checked);" value="" title="Toggle on sale checkboxes"></th>
  </tr>
</table>

<script language="javascript">
<!--

function makeDirty(lngID, blnAttribute)
{

	var plngCount;
	var i;

	if (blnAttribute)
	{
		plngCount = document.frmData.attrdtID.length;
		if (document.frmData.attrdtID.checked==undefined)
		{
			for (i=0; i < plngCount;i++)
			{
				if (document.frmData.attrdtID[i].value == lngID)
				{
				document.frmData.attrdtID[i].checked = true;
				return true
				}
			}
		}
	}else{
		plngCount = document.frmData.prodID.length;
		if (document.frmData.prodID.checked==undefined)
		{
			for (i=0; i < plngCount;i++)
			{
				if (document.frmData.prodID[i].value == lngID)
				{
				document.frmData.prodID[i].checked = true;
				return true
				}
			}
		}
	}

}

//-->
</script>
<script language="vbscript">

Function getElementName(theElement)

Dim pstrName

	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	pstrName = theElement.name
	If Err.number <> 0 Then Err.Clear
	
	getElementName = pstrName
End Function

Function SetPrices(theForm)

Dim i
Dim theElement
Dim pstrName
Dim plngPos
Dim pdblOrigPrice
Dim pdblNewPrice

Dim pblnIncrease
Dim pblnPercent
Dim pdblIncreaseBy
Dim pdblEndWith

	pblnIncrease = theForm.radSetPrice(0).checked
	pblnPercent = theForm.radPriceIncreaseBy(0).checked
	pdblIncreaseBy = Trim(theForm.txtPriceIncreaseBy.value)
	pdblEndWith = Trim(theForm.txtEndPriceIncrease.value)

	If (pdblIncreaseBy = "") Then
		msgBox("Please enter an amount to increase/decrease prices by")
		theForm.txtPriceIncreaseBy.focus()
		Exit Function
	End If

	If Not isNumeric(pdblIncreaseBy) Then
		msgBox("Please enter a number for the increase/decrease prices by")
		theForm.txtPriceIncreaseBy.select()
		theForm.txtPriceIncreaseBy.focus()
		Exit Function
	End If

	If pdblEndWith <> "" And Not isNumeric(pdblEndWith) Then
		msgBox("Please enter a number to end the prices by")
		theForm.txtEndPriceIncrease.select()
		theForm.txtEndPriceIncrease.focus()
		Exit Function
	ElseIf pdblEndWith <> "" Then
		If CLng(pdblEndWith) < 0 OR CLng(pdblEndWith) > 99 Then
			msgBox("Please enter a number between 00 and 99 to end the prices by")
			theForm.txtEndPriceIncrease.select()
			theForm.txtEndPriceIncrease.focus()
			Exit Function
		End If
	End If

	For i = 0 To theForm.length - 1
		Set theElement = theForm.elements(i)
		pstrName = getElementName(theElement)
		plngPos = inStr(1, pstrName, "prodPrice")
		If (plngPos > 0) Then
			pdblOrigPrice = theElement.value
			pdblNewPrice = AlterPrice(pdblOrigPrice, pblnIncrease, pblnPercent, pdblIncreaseBy, pdblEndWith)
			theElement.value = pdblNewPrice
		End If
	Next 'i

End Function	'SetPrices

'***************************************************************************************************************************************************

Function SetSalePrices(theForm)

Dim i
Dim theElement
Dim pstrName
Dim plngPos
Dim pdblOrigPrice
Dim pdblNewPrice

Dim pblnIncrease
Dim pblnPercent
Dim pdblIncreaseBy
Dim pdblEndWith
Dim pblnPriceBasedOnRegularPrice
Dim pstrProdID

	pblnIncrease = theForm.radSetSalePrice(0).checked
	pblnPercent = theForm.radSalePriceIncreaseBy(0).checked
	pblnPriceBasedOnRegularPrice = theForm.radSalePriceSource(1).checked
	pdblIncreaseBy = Trim(theForm.txtSalePriceIncreaseBy.value)
	pdblEndWith = Trim(theForm.txtEndSalePriceIncrease.value)

	If (pdblIncreaseBy = "") Then
		msgBox("Please enter an amount to increase/decrease sale prices by")
		theForm.txtSalePriceIncreaseBy.focus()
		Exit Function
	End If

	If Not isNumeric(pdblIncreaseBy) Then
		msgBox("Please enter a number for the increase/decrease sale prices by")
		theForm.txtSalePriceIncreaseBy.select()
		theForm.txtSalePriceIncreaseBy.focus()
		Exit Function
	End If

	If pdblEndWith <> "" And Not isNumeric(pdblEndWith) Then
		msgBox("Please enter a number to end the sale prices by")
		theForm.txtEndSalePriceIncrease.select()
		theForm.txtEndSalePriceIncrease.focus()
		Exit Function
	ElseIf pdblEndWith <> "" Then
		If CLng(pdblEndWith) < 0 OR CLng(pdblEndWith) > 99 Then
			msgBox("Please enter a number between 00 and 99 to end the sale prices by")
			theForm.txtEndSalePriceIncrease.select()
			theForm.txtEndSalePriceIncrease.focus()
			Exit Function
		End If
	End If

	For i = 0 To theForm.length - 1
		Set theElement = theForm.elements(i)
		pstrName = getElementName(theElement)
		plngPos = inStr(1, pstrName, "prodSalePrice")
		If (plngPos > 0) Then
			pstrProdID = Replace(pstrName,"prodSalePrice","")
			If pblnPriceBasedOnRegularPrice Then
				pdblOrigPrice = document.all("prodPrice" & pstrProdID).value
			Else
				pdblOrigPrice = theElement.value
			End If

			pdblNewPrice = AlterPrice(pdblOrigPrice, pblnIncrease, pblnPercent, pdblIncreaseBy, pdblEndWith)
			theElement.value = pdblNewPrice
		End If
	Next 'i

End Function	'SetSalePrices

'***************************************************************************************************************************************************

Function SetAttrPrices(theForm)

Dim i
Dim theElement
Dim pstrName
Dim plngPos
Dim pdblOrigPrice
Dim pdblNewPrice

Dim pblnIncrease
Dim pblnPercent
Dim pdblIncreaseBy
Dim pdblEndWith
Dim pblnPriceBasedOnRegularPrice
Dim pstrProdID

	pblnIncrease = theForm.radSetAttrPrice(0).checked
	pblnPercent = theForm.radAttrPriceIncreaseBy(0).checked
	pblnPriceBasedOnRegularPrice = theForm.radAttrPriceSource(1).checked
	pdblIncreaseBy = Trim(theForm.txtAttrPriceIncreaseBy.value)
	pdblEndWith = Trim(theForm.txtEndAttrPriceIncrease.value)

	If (pdblIncreaseBy = "") Then
		msgBox("Please enter an amount to increase/decrease attribute prices by")
		theForm.txtAttrPriceIncreaseBy.focus()
		Exit Function
	End If

	If Not isNumeric(pdblIncreaseBy) Then
		msgBox("Please enter a number for the increase/decrease attribute prices by")
		theForm.txtAttrPriceIncreaseBy.select()
		theForm.txtAttrPriceIncreaseBy.focus()
		Exit Function
	End If

	If pdblEndWith <> "" And Not isNumeric(pdblEndWith) Then
		msgBox("Please enter a number to end the sale prices by")
		theForm.txtEndAttrPriceIncrease.select()
		theForm.txtEndAttrPriceIncrease.focus()
		Exit Function
	ElseIf pdblEndWith <> "" Then
		If CLng(pdblEndWith) < 0 OR CLng(pdblEndWith) > 99 Then
			msgBox("Please enter a number between 00 and 99 to end attribute sale prices by")
			theForm.txtEndAttrPriceIncrease.select()
			theForm.txtEndAttrPriceIncrease.focus()
			Exit Function
		End If
	End If

	For i = 0 To theForm.length - 1
		Set theElement = theForm.elements(i)
		pstrName = getElementName(theElement)
		plngPos = inStr(1, pstrName, "attrdtPrice")
		If (plngPos > 0) Then
			pstrProdID = theElement.tag
			If pblnPriceBasedOnRegularPrice Then
				pdblOrigPrice = document.all("prodPrice" & pstrProdID).value
			Else
				pdblOrigPrice = theElement.value
			End If

			pdblNewPrice = AlterPrice(pdblOrigPrice, pblnIncrease, pblnPercent, pdblIncreaseBy, pdblEndWith)
			theElement.value = pdblNewPrice
		End If
	Next 'i

End Function	'SetAttrPrices

'***************************************************************************************************************************************************

Function AlterPrice(dblBasePrice, blnIncrease, blnPercent, dblIncreaseBy, dblEndWith)

Dim pdblPrice
Dim pdblPriceChange
Dim pdblTempPrice

	If (blnPercent) Then
		pdblPriceChange = CDbl(dblIncreaseBy)/100 * CDbl(dblBasePrice)
	Else
		pdblPriceChange = CDbl(dblIncreaseBy)
	End If

	If (blnIncrease) Then
		pdblPrice = CDbl(dblBasePrice) + CDbl(pdblPriceChange)
	Else
		pdblPrice = CDbl(dblBasePrice) - CDbl(pdblPriceChange)
	End If

	If (dblEndWith <> "") Then
		pdblTempPrice = Int(pdblPrice) + CInt(dblEndWith)/100
		If CDbl(pdblTempPrice) < CDbl(pdblPrice) Then
			pdblPrice = pdblTempPrice + 1
		Else
			pdblPrice = pdblTempPrice
		End If		
	End If
	
	AlterPrice = Round(pdblPrice,2)
	
End Function	'AlterPrice

</script>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblPriceChange">
  <tr><td><hr width="90%"</td></tr>
  <tr>
    <td valign="top" align="center">
    <h4>Adjust Regular Prices</h4>
    <table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="Table1">
      <tr>
        <td>
        <input type="radio" name="radSetPrice" value="0" id="radSetPrice1" checked><label for="radSetPrice1">Increase</label><br />
        <input type="radio" name="radSetPrice" value="1" id="radSetPrice2"><label for="radSetPrice2">Decrease</label></td>
        <td>regular prices by</td>
        <td><input type="text" name="txtPriceIncreaseBy" size="4" id="Text1">
        </td>
        <td>
        <input type="radio" name="radPriceIncreaseBy" value="0" id="radPriceIncreaseBy1"><label for="radPriceIncreaseBy1">Percent</label><br />
        <input type="radio" name="radPriceIncreaseBy" value="1" id="radPriceIncreaseBy2" checked><label for="radPriceIncreaseBy2">Amount</label></td>
        <td>end price with .<input type="text" name="txtEndPriceIncrease" title="Set this field to the decimal you wish to end the price with. The price will always be rounded up to this number. Leave blank if you do not want to force this number." size="2" maxlength="2" id="Text2"></td>
        <td align="center">
        <input class="butn" title="Save changes" id="btnSetPrices" name="btnSetPrices" type="button" value="Set Prices" onclick="SetPrices(this.form);"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr><td><hr width="90%"</td></tr>
  <tr>
    <td valign="top" align="center">
    <h4>Adjust Sale Prices</h4>
    <table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="Table2">
      <tr>
        <td>
        <input type="radio" name="radSetSalePrice" value="0" id="radSetSalePrice1" checked><label for="radSetSalePrice1">Increase</label><br />
        <input type="radio" name="radSetSalePrice" value="1" id="radSetSalePrice2"><label for="radSetSalePrice2">Decrease</label></td>
        <td>sale prices by</td>
        <td><input type="text" name="txtSalePriceIncreaseBy" size="4" id="txtSalePriceIncreaseBy">
        </td>
        <td>
        <input type="radio" name="radSalePriceIncreaseBy" value="0" id="radSalePriceIncreaseBy1"><label for="radSalePriceIncreaseBy1">Percent</label><br />
        <input type="radio" name="radSalePriceIncreaseBy" value="1" id="radSalePriceIncreaseBy2" checked><label for="radSalePriceIncreaseBy2">Amount</label></td>
        <td>end price with .<input type="text" name="txtEndSalePriceIncrease" id="txtEndSalePriceIncrease" title="Set this field to the decimal you wish to end the price with. The price will always be rounded up to this number. Leave blank if you do not want to force this number." size="2" maxlength="2"></td>
        <td>Base source price on</td>
        <td>
        <input type="radio" name="radSalePriceSource" value="0" id="radSalePriceSource1"><label for="radSalePriceSource1">Existing Sale Price</label><br />
        <input type="radio" name="radSalePriceSource" value="1" id="radSalePriceSource2" checked><label for="radSalePriceSource2">Regular Price</label></td>
        <td align="center">
        <input class="butn" title="Save changes" id="btnSetSalePrices" name="btnSetSalePrices" type="button" value="Set Sale Prices" onclick="SetSalePrices(this.form);"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr><td><hr width="90%"</td></tr>
  <tr>
    <td valign="top" align="center">
    <h4>Adjust Attribute Prices</h4>
    <table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="Table3">
      <tr>
        <td>
        <input type="radio" name="radSetAttrPrice" value="0" id="radSetAttrPrice1" checked><label for="radSetAttrPrice1">Increase</label><br />
        <input type="radio" name="radSetAttrPrice" value="1" id="radSetAttrPrice2"><label for="radSetAttrPrice2">Decrease</label></td>
        <td>attribute prices by</td>
        <td><input type="text" name="txtAttrPriceIncreaseBy" size="4" id="txtAttrPriceIncreaseBy">
        </td>
        <td>
        <input type="radio" name="radAttrPriceIncreaseBy" value="0" id="radAttrPriceIncreaseBy1"><label for="radAttrPriceIncreaseBy1">Percent</label><br />
        <input type="radio" name="radAttrPriceIncreaseBy" value="1" id="radAttrPriceIncreaseBy2" checked><label for="radAttrPriceIncreaseBy2">Amount</label></td>
        <td>end price with .<input type="text" name="txtEndAttrPriceIncrease" id="txtEndAttrPriceIncrease" title="Set this field to the decimal you wish to end the price with. The price will always be rounded up to this number. Leave blank if you do not want to force this number." size="2" maxlength="2"></td>
        <td>Base source price on</td>
        <td>
        <input type="radio" name="radAttrPriceSource" value="0" id="radAttrPriceSource1" checked><label for="radAttrPriceSource1">Existing Attribute Price</label><br />
        <input type="radio" name="radAttrPriceSource" value="1" id="radAttrPriceSource2"><label for="radAttrPriceSource2">Regular Price</label></td>
        <td align="center">
        <input class="butn" title="Save changes" id="btnSetAttrPrices" name="btnSetAttrPrices" type="button" value="Set Attribute Prices" onclick="SetAttrPrices(this.form);"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr><td><hr width="90%"</td></tr>
</table>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
		<input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/ProductPricingTool_help.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
        <input class="butn" title="Save changes" id="btnUpdate" name="btnUpdate" type="submit" value="Save Changes">
    </td>
  </tr>
</table>

</form>

</center>
</BODY>
</HTML>
<%

	Call ReleaseObject(mrsCategory)
	Call ReleaseObject(mrsVendor)
	Call ReleaseObject(mrsManufacturer)
	
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
