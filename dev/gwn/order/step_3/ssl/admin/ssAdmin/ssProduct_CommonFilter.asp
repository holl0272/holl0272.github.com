<%
'********************************************************************************
'*   Product Manager Common Filter Version SF 5.0 		                        *
'*   Release Version   1.01				                                        *
'*   Release Date      October 12, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'***********************************************************************************************

Function AECategoryFilter(ByVal bytCategoryFilter, ByVal bytsubCategoryFilter)

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
			Dim plngCatID
			Dim pstrCatValue
			Dim pstrCatName
			Dim pbytCatDepth
			Dim pstrsubCatValue
			Dim pstrsubCatName
			Dim pstrSelected
			Dim pblnBottom
			
			Dim pstrOptionValue

			If Len(bytCategoryFilter) = 0 And Len(bytsubCategoryFilter) = 0 Then pstrSelected = "selected"
			pstrTemp = "<select id='CategoryFilter' name='CategoryFilter' size='10' multiple>" & vbcrlf _
					 & "  <option value='...' " & pstrSelected & ">- All -</Option>" & vbcrlf

			For i=1 to .RecordCount
				If len(.Fields("Depth").value & "") > 0 Then 'Then No sub-categories so don't display
					pbytCatDepth = .Fields("Depth").value
					
					If plngCatID <> Trim(.Fields("catID").value) Then
						plngCatID = Trim(.Fields("catID").value)
						
						pstrCatValue = .Fields("catID").value & "."
						pstrCatName = .Fields("catName").value
						pstrOptionValue = pstrCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value
						
						pstrSelected = ""
						If pstrCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
						pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrCatName & "</option>" & vbcrlf 
						
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrOptionValue = pstrsubCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value

							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					Else
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrOptionValue = pstrsubCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value
							
							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					End If
				End If
				
				.MoveNext
			Next
			pstrTemp = pstrTemp & "</select>" & vbcrlf
			
			
		Else
			pstrTemp = "<font color=red>No Categories</font><br>"
			pstrTemp = pstrTemp & "<select id=CatSource name=CatSource size=10 multiple>"
			pstrTemp = pstrTemp & "</select>"
		End If
	End With

	AECategoryFilter = pstrTemp
	
End Function	'AECategoryFilter

'******************************************************************************************************************************************************************

Sub LoadFilter

dim pstrsqlWhere

	'modified so could link in directly
	mlngManufacturerFilter = LoadRequestValue("ManufacturerFilter")
	mlngCategoryFilter = LoadRequestValue("CategoryFilter")
	mlngVendorFilter = LoadRequestValue("VendorFilter")
	mradTextSearch = LoadRequestValue("radTextSearch")
	mstrTextSearch = trim(LoadRequestValue("TextSearch"))
	mcurUpperPrice = trim(LoadRequestValue("UpperPrice"))
	mcurLowerPrice = trim(LoadRequestValue("LowerPrice"))
	mradShowActive = LoadRequestValue("radShowActive")
	mradShowOnsale = LoadRequestValue("radShowOnsale")
	mradShowShipped = LoadRequestValue("radShowShipped")

	'Now check for subCategories
	Dim paryCat,paryTempCat
	Dim pstrTempCat, pstrTempSubCat
	Dim pbytTempDepth, pblnTempBottom
	
	paryTempCat = Split(mlngCategoryFilter,",")
	If Len(mlngCategoryFilter) > 0 Then
		If isArray(paryTempCat) Then
			paryCat = Split(paryTempCat(0),".")
			pstrTempCat = Trim(paryCat(0))
			pstrTempSubCat = Trim(paryCat(1))
			pbytTempDepth = Trim(paryCat(2))
			pblnTempBottom = Trim(paryCat(3))
			For i = 1 To UBound(paryTempCat)
				paryCat = Split(paryTempCat(i),".")
				pstrTempCat = pstrTempCat & "," & Trim(paryCat(0))
				pstrTempSubCat = pstrTempSubCat & "," & Trim(paryCat(1))
				pbytTempDepth = pbytTempDepth & "," & Trim(paryCat(2))
				pblnTempBottom = pblnTempBottom & "," & Trim(paryCat(3))
			Next 'i
		
			mbytCategoryFilter = pstrTempCat
			mbytSubCategoryFilter = pstrTempSubCat
					
			'debugprint "mbytCategoryFilter",mbytCategoryFilter
			'debugprint "mbytSubCategoryFilter",mbytSubCategoryFilter
			'Response.Flush
		Else
			mbytCategoryFilter = mlngCategoryFilter
		End If
	End If
	
	If Not isNumeric(mcurUpperPrice) Then mcurUpperPrice = ""
	If Not isNumeric(mcurLowerPrice) Then mcurLowerPrice = ""

	If len(mstrTextSearch) > 0 Then
		Select Case mradTextSearch
			Case "0"	'Do Not Include
			Case "1"	'ProductID
				pstrsqlWhere = "sfProducts.prodID Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "2"	'Short Description
				pstrsqlWhere = "sfProducts.prodShortDescription Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "3"	'Long Description
				pstrsqlWhere = "sfProducts.prodDescription Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "4"	'Product Name
				pstrsqlWhere = "sfProducts.prodName Like '%" & sqlSafe(mstrTextSearch) & "%'"
		End Select	
	End If

	Select Case mradShowActive
		Case "0"	'Do Not Include
		Case "1"	'Active
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodEnabledIsActive=1"
			Else
				pstrsqlWhere = "prodEnabledIsActive=1"
			End If
		Case "2"	'Inactive
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodEnabledIsActive=0"
			Else
				pstrsqlWhere = "prodEnabledIsActive=0"
			End If
	End Select	
	
	Select Case mradShowShipped
		Case "0"	'Do Not Include
		Case "1"	'Shipped
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodShipIsActive=1"
			Else
				pstrsqlWhere = "prodShipIsActive=1"
			End If
		Case "2"	'Not Shipped
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodShipIsActive<>1"
			Else
				pstrsqlWhere = "prodShipIsActive<>1"
			End If
	End Select	
	
	Select Case mradShowOnsale
		Case "0"	'Do Not Include
		Case "1"	'Regular
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodSaleIsActive<>1"
			Else
				pstrsqlWhere = "prodSaleIsActive<>1"
			End If
		Case "2"	'Onsale
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodSaleIsActive=1"
			Else
				pstrsqlWhere = "prodSaleIsActive=1"
			End If
	End Select	
	
	If len(mlngManufacturerFilter) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and prodManufacturerId=" & mlngManufacturerFilter
		Else
			pstrsqlWhere = "prodManufacturerId=" & mlngManufacturerFilter
		End If
	End If


	If len(mbytCategoryFilter) > 0 then

		If cblnSF5AE Then	' And False
			Dim paryCategory, parysubCategory
			Dim paryDepth, paryBottom
			Dim i
			Dim pstrTemp
			Dim pstrCatSubCatChoice
			
			paryCategory = Split(mbytCategoryFilter,",")
			If Len(mbytsubCategoryFilter) = 0 Then
				parysubCategory = Array("")
			Else
				parysubCategory = Split(mbytsubCategoryFilter,",")
			End If

			If Len(pbytTempDepth) = 0 Then
				paryDepth = Array("")
			Else
				paryDepth = Split(pbytTempDepth,",")
			End If

			If Len(pblnTempBottom) = 0 Then
				paryBottom = Array("")
			Else
				paryBottom = Split(pblnTempBottom,",")
			End If

			'paryCategory = Split(mbytCategoryFilter,",")
			If isArray(paryCategory) Then
				For i = 0 To UBound(paryCategory)
					If Len(parysubCategory(i)) > 0 Then
						If paryBottom(i) = 1 Then
							pstrCatSubCatChoice = "(sfSub_Categories.subcatId=" & parysubCategory(i) & ")"
						Else
							If paryDepth(i) = 1 Then
								pstrCatSubCatChoice = "(sfSub_Categories.CatHierarchy Like '" & parysubCategory(i) & "-%')"
							Else
								pstrCatSubCatChoice = "(sfSub_Categories.CatHierarchy Like '%-" & parysubCategory(i) & "-%')"
							End If
						End If
						'Response.Write paryCategory(i) & "." & paryDepth(i) & " - " & paryBottom(i) & "<BR>"
					ElseIf Len(paryCategory(i)) > 0 Then
						pstrCatSubCatChoice = "(sfSub_Categories.subcatCategoryId=" & paryCategory(i) & ")"
					Else
						pstrCatSubCatChoice = ""
					End If
					
					If Len(pstrTemp) > 0 Then
						pstrTemp = pstrTemp & " OR " & pstrCatSubCatChoice
					Else
						pstrTemp = "( " & pstrCatSubCatChoice
					End If
				Next 'i
				If Request.Form("radCategoryInclude") <> "OR" Then mstrsqlHaving = " HAVING (Count(sfProducts.prodName)=" & CStr(UBound(paryCategory) + 1) & ")"
			Else
				pstrsqlWhere = "sfSub_Categories.subcatCategoryId=" & mbytCategoryFilter & ""
			End If
			pstrTemp = pstrTemp & ")"

		Else
			pstrTemp = "prodCategoryId=" & mbytCategoryFilter
		End If
		'debugprint "mstrsqlHaving",mstrsqlHaving
		'debugprint "pstrTemp",pstrTemp

		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " AND " & pstrTemp
		Else
			pstrsqlWhere = pstrTemp
		End If
	End If

	If len(mlngVendorFilter) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and prodVendorId=" & mlngVendorFilter
		Else
			pstrsqlWhere = "prodVendorId=" & mlngVendorFilter
		End If
	End If

	If len(mcurLowerPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,prodPrice)>=" & mcurLowerPrice
			Else
				pstrsqlWhere = "convert(money,prodPrice)>=" & mcurLowerPrice
			End If
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and cCur(prodPrice)>=" & mcurLowerPrice
			Else
				pstrsqlWhere = "cCur(prodPrice)>=" & mcurLowerPrice
			End If
		End If
	End If

	If len(mcurUpperPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,prodPrice)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "convert(money,prodPrice)<=" & mcurUpperPrice
			End If
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and cCur(prodPrice)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "cCur(prodPrice)<=" & mcurUpperPrice
			End If
		End If
	End If

	If len(pstrsqlWhere) > 0 then mstrsqlWhere = "where "  & pstrsqlWhere

End Sub    'LoadFilter

'******************************************************************************************************************************************************************

Sub LoadSort

Dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Product ID
			pstrOrderBy = "sfProducts.prodID"
		Case "2"	'Product Name
			pstrOrderBy = "prodName"
		Case "3"	'Price
			If cblnSQLDatabase Then
				pstrOrderBy = "convert(money,prodPrice)"
			Else
				pstrOrderBy = "cCur(prodPrice)"
			End If
		Case "4"	'Active
			pstrOrderBy = "prodEnabledIsActive"
	End Select	

	If len(pstrOrderBy) > 0 then mstrsqlWhere = mstrsqlWhere  & " Order By " & pstrOrderBy & " " & mstrSortOrder
	
End Sub    'LoadSort

'******************************************************************************************************************************************************************

Dim mstrsqlWhere, mstrsqlHaving, mstrSortOrder,mstrOrderBy

Dim mlngManufacturerFilter,mlngCategoryFilter,mlngVendorFilter
Dim mbytCategoryFilter, mbytSubCategoryFilter
Dim mradTextSearch, mstrTextSearch, mcurUpperPrice, mcurLowerPrice, mradShowActive, mradShowOnsale, mradShowShipped
Dim mblnDetailInNewWindow
Dim mblnAutoShowDetailInWindow

	Dim mstrTemp
	mstrTemp = LoadRequestValue("chkDetailInNewWindow")

	If Instr(1,mstrTemp,",") > 0 Then
		mblnDetailInNewWindow = Trim(Right(mstrTemp,Len(mstrTemp)-Instr(1,mstrTemp,",")))
	Else
		mblnDetailInNewWindow = mstrTemp
	End If
	If Len(mblnDetailInNewWindow) = 0 Then 
		If mblnAutoShowDetailInWindow Then
			mblnDetailInNewWindow = 0
		Else
			mblnDetailInNewWindow = 1
		End If
	End If
	
	
Sub WriteProductFilter(byVal blnShowNewWindowChoice)
%>
<SCRIPT LANGUAGE=javascript>
<!--

function btnFilter_onclick(theButton)
{
var theForm = theButton.form;

  if (!isNumeric(theForm.LowerPrice,true,"Please enter a number for the price.")) {return(false);}
  if (!isNumeric(theForm.UpperPrice,true,"Please enter a number for the price.")) {return(false);}

  theForm.Action.value = "Filter";
  theForm.submit();
  return(true);
}

//-->
</SCRIPT>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
    <td valign="top">
        <input type="radio" value="1" <% if mradTextSearch="1" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch1"><label for="radTextSearch1">Product ID</label><br>
        <input type="radio" value="2" <% if mradTextSearch="2" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch2"><label for="radTextSearch2">Short Description</label><br>
        <input type="radio" value="3" <% if mradTextSearch="3" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch3"><label for="radTextSearch3">Long Description</label><br>
        <input type="radio" value="4" <% if mradTextSearch="4" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch4"><label for="radTextSearch4">Product Name</label><br>
        <input type="radio" value="0" <% if (mradTextSearch="0" or mradTextSearch="") then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch0"><label for="radTextSearch0">Do Not Include</label>
        <p>containing the text<br>
        <input type="text" name="TextSearch" size="20" value="<%= EncodeString(mstrTextSearch,True) %>" ID="TextSearch">
	</td>
	<td valign="top" align="center">
		Price between<br>
        <p><input type="text" id="LowerPrice" name="LowerPrice" size="10" value="<%= mcurLowerPrice %>" maxlength="15"></p>
        <p>And</p>
        <input type="text" id="UpperPrice" name="UpperPrice" size="10" value="<%= mcurUpperPrice %>" maxlength="15">
	</td>
	<td valign="top" align="center">
        Filter by Manufacturer<br>
	<select size="1"  id=ManufacturerFilter name=ManufacturerFilter>
<% 	if len(mlngManufacturerFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	Call MakeCombo(mrsManufacturer,"mfgName","mfgID",mlngManufacturerFilter)
 %>
	</select>
	<p>Filter by Category<br>
<% If cblnSF5AE Then 
		Response.Write AECategoryFilter(mbytCategoryFilter, mbytsubCategoryFilter)
   Else %>
	<select size="1"  id="CategoryFilter" name="CategoryFilter">
	<% 	if len(mlngCategoryFilter) = 0 then
			Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
		else
			Response.Write "<option value=''>- All -</Option>" & vbcrlf
		end if
		Call MakeCombo(mrsCategory,"catName","catID",mlngCategoryFilter)
	%>
	</select>
<% End If %>

<% If cblnSF5AE Then %>
<br>
<input type="radio" value="OR" name="radCategoryInclude" id="radCategoryInclude1" <% If Request.Form("radCategoryInclude") <> "AND" Then Response.Write "checked" %>><label for="radCategoryInclude1">OR</label>&nbsp;
<input type="radio" value="AND" name="radCategoryInclude" id="radCategoryInclude2" <% If Request.Form("radCategoryInclude") = "AND" Then Response.Write "checked" %>><label for="radCategoryInclude2">AND</label>
<% End If %>
	<p>Filter by Vendor<br>
	<select size="1"  id=VendorFilter name=VendorFilter>
<% 	if len(mlngVendorFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	Call MakeCombo(mrsVendor,"vendName","vendID",mlngVendorFilter)
%>
	</select>
	</td>
    <td valign="top">
        <p>Show Products that are:<br>
        <input type="radio" value="1" <% if mradShowActive="1" then Response.Write "Checked" %> name="radShowActive" ID="radShowActive1"><label for="radShowActive1">Active</label>&nbsp;
        <input type="radio" value="2" <% if mradShowActive="2" then Response.Write "Checked" %> name="radShowActive" ID="radShowActive2"><label for="radShowActive2">Inactive</label>&nbsp;
        <input type="radio" value="0" <% if (mradShowActive="0" or mradShowActive="") then Response.Write "Checked" %> name="radShowActive" ID="radShowActive0"><label for="radShowActive0">All</label><br>

        <input type="radio" value="1" <% if mradShowOnsale="1" then Response.Write "Checked" %> name="radShowOnsale" ID="radShowOnsale1"><label for="radShowOnsale1">Regular</label>&nbsp;
        <input type="radio" value="2" <% if mradShowOnsale="2" then Response.Write "Checked" %> name="radShowOnsale" ID="radShowOnsale2"><label for="radShowOnsale2">On Sale</label>&nbsp;
        <input type="radio" value="0" <% if (mradShowOnsale="0" or mradShowOnsale="") then Response.Write "Checked" %> name="radShowOnsale" ID="radShowOnsale0"><label for="radShowOnsale0">All</label><br>

        <input type="radio" value="1" <% if mradShowShipped="1" then Response.Write "Checked" %> name="radShowShipped" ID="radShowShipped1"><label for="radShowShipped1">Shipped</label>&nbsp;
        <input type="radio" value="2" <% if mradShowShipped="2" then Response.Write "Checked" %> name="radShowShipped" ID="radShowShipped2"><label for="radShowShipped2">Not Shipped</label>&nbsp;
        <input type="radio" value="0" <% if (mradShowShipped="0" or mradShowShipped="") then Response.Write "Checked" %> name="radShowShipped" ID="radShowShipped0"><label for="radShowShipped0">All</label><br>
	</td>
	<td>
	  <input class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"><br>
	  <% If blnShowNewWindowChoice Then %>
	  <input type="radio" name="chkDetailInNewWindow" id="chkDetailInNewWindow0" value="0" <% if mblnDetailInNewWindow=0 then Response.Write "Checked" %>>&nbsp;<label for="chkDetailInNewWindow0">Open detail in this window</label><br>
	  <input type="radio" name="chkDetailInNewWindow" id="chkDetailInNewWindow1" value="1" <% if mblnDetailInNewWindow=1 then Response.Write "Checked" %>>&nbsp;<label for="chkDetailInNewWindow1">Open detail in new window</label>
	  <% End If %>
	</td>
  </tr>
</table>
<% End Sub	'WriteProductFilter %>