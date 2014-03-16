<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: November 26, 2004											*
'*   Version 1.01.001                                                           *
'*                                                                              *
'*   1.00.001 (November 26, 2004)                                               *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************

'***********************************************************************************************

'Array decoder
' 0- sfProducts.ProdID
' 1- sfProducts.prodName
' 2- sfProducts.prodImageSmallPath
' 3- sfProducts.prodLink
' 4- sfProducts.prodPrice
' 5- sfProducts.prodSaleIsActive
' 6- sfProducts.prodSalePrice
' 7- sfProducts.catName
' 8- sfProducts.prodDescription
' 9- sfProducts.prodAttrNum
' 10- sfProducts.prodCategoryId
' 11- sfProducts.prodShortDescription
' 12- prodPLPrice</H3>
' 13- prodPLSalePrice</H3>
' 14- catDescription</H3>
' 15- catImage
' 16- vendName
' 17- mfgName

'**********************************************************
'*	Page Level variables
'**********************************************************

Const enPageType_Search = 0
Const enPageType_Category = 1
Const enPageType_Mfg = 2
Const enPageType_Vend = 3
Const enPageType_Sales = 4
Const enPageType_NewProduct = 5
	
Dim mbytDisplayStyle
Dim mbytPageType
Dim mbytSortBy
Dim mstrHideImages
Dim mstrSortOrder
Dim mstrCustomDescription
Dim mstrCustomImage
Dim mstrCustomName
Dim mstrCustomNameWrapped

Dim mblnCategorySearch
Dim mblnSubCategorySearch
Dim mblnManufacturerSearch
Dim mblnVendorSearch
Dim mstrProductIDList
Dim mlngNumProductsFound

'**********************************************************
'*	Functions
'**********************************************************

'Function BuildCustomSearchFilter(byRef strSQL, byVal searchParamType, byVal searchParamTxt, byVal searchParamCat, byVal searchParamMan, byVal searchParamVen, byVal DateAddedStart, byVal DateAddedEnd, byVal PriceStart, byVal PriceEnd, byVal Sale, byVal subCatID, byVal Ilevel)
'Sub LoadSearchParameters
'Sub LoadSort(byRef strSQL, byRef lngPageSize)

'**********************************************************
'*	Begin Page Code
'**********************************************************

Function categorySearchWhere(byVal subCatID, byVal Ilevel)

Dim pstrSQL_Where

	mblnSubCategorySearch = CBool(Ilevel > 1)
	If CBool(Ilevel < 2) And Len(subCatID) > 0 And isNumeric(subCatID) Then
		pstrSQL_Where = "(sfSubCatDetail.subcatCategoryId In (Select subcatID From sfSub_Categories Where subcatCategoryId= " & subCatID & "))"
	ElseIf Len(subCatID) > 0 And isNumeric(subCatID) Then
		pstrSQL_Where = "(" _
						& "  (sfSubCatDetail.subcatCategoryId= " & subCatID & ")" _
						& " OR (sfSubCatDetail.subcatCategoryId In (Select subcatID From sfSub_Categories Where left(CatHierarchy," & Len(CStr(subCatID)) + 1 & ")='" & subCatID & "-'))" _
						& ")"
	End If

	If vDebug = 1 Then
		Response.Write "<fieldset><legend>categorySearchWhere</legend>"
		Response.Write "Ilevel: " & Ilevel & "<br /><hr>"
		Response.Write "subCatID: " & subCatID & "<br /><hr>"
		Response.Write "pstrSQL_Where: " & pstrSQL_Where & "<br /><hr>"
		Response.Write "</fieldset>"
	End If
	
	categorySearchWhere = pstrSQL_Where
	
End Function	'categorySearchWhere

'**********************************************************

Function getProductIDsMatchingSearch(byVal searchParamType, byVal searchParamTxt, byVal searchParamCat, byVal searchParamMan, byVal searchParamVen, byVal DateAddedStart, byVal DateAddedEnd, byVal PriceStart, byVal PriceEnd, byVal Sale, byVal subCatID, byVal Ilevel)

Dim counter
Dim pstrSQL
Dim pstrSQL_GroupBy
Dim pstrSQL_Where
Dim txtArray
Dim upperLim
Dim pobjRS

	If cblnSF5AE Then
	
		If cblnSearchAttributes Then
			pstrSQL = "SELECT sfProducts.prodID" _
					& " FROM ((sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId"
		Else
			pstrSQL = "SELECT sfProducts.prodID" _
					& " FROM sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID"
		End If

		pstrSQL_Where = categorySearchWhere(subCatID, Ilevel)
		If Len(pstrSQL_Where) > 0 Then
			pstrSQL_Where = " WHERE (sfProducts.prodEnabledIsActive=1) And " & pstrSQL_Where
		Else
			pstrSQL_Where = " WHERE (sfProducts.prodEnabledIsActive=1)"
		End If
		
		If cblnSQLDatabase Then
			'Cannot group by description since it is ntext
			pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodShortDescription, sfProducts.prodNamePlural" & SearchResults_OrderBy(True)
		Else
			pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodShortDescription, sfProducts.prodNamePlural" & SearchResults_OrderBy(True)
		End If
		
	Else
		If cblnSearchAttributes Then
			pstrSQL = "SELECT sfProducts.prodID" _
					& " FROM sfAttributeDetail RIGHT JOIN (sfAttributes RIGHT JOIN (sfManufacturers RIGHT JOIN (sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfAttributes.attrProdId = sfProducts.prodID) ON sfAttributeDetail.attrdtAttributeId = sfAttributes.attrID"
		Else
			pstrSQL = "SELECT sfProducts.prodID" _
					& " FROM sfManufacturers RIGHT JOIN (sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerID"
		End If

		If iSubCat = "ALL" Then
			pstrSQL_Where = " WHERE "
		Else
			If Len(searchParamCat) > 0 And isNumeric(searchParamCat) Then
				pstrSQL_Where = " WHERE sfProducts.prodEnabledIsActive=1 And sfProducts.prodCategoryId=" & searchParamCat
			Else
				pstrSQL_Where = " WHERE sfProducts.prodEnabledIsActive=1"
			End If
		End If 
	End If	'cblnSF5AE

    If Len(searchParamMan) > 0 And isNumeric(searchParamMan) Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts.prodManufacturerId=" & wrapSQLValue(searchParamMan, False, enDatatype_number)
    If Len(searchParamVen) > 0 And isNumeric(searchParamVen) Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts.prodVendorId=" & wrapSQLValue(searchParamVen, False, enDatatype_number)

	If Len(searchParamTxt) > 0 Then 
		If searchParamType = "ALL" Then 
			txtArray = split(searchParamTxt, " ")
			
			If Len(searchParamTxt) > 0 Then
				For counter = 0 to Ubound(txtArray)
					Select Case Request.QueryString("narrowSearch")
						Case "prodID"
							pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodID LIKE '%" & txtArray(counter) & "%') "
						Case "mfgID"
							pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodName LIKE '%" & txtArray(counter) & "%') "
						Case Else
							pstrSQL_Where = pstrSQL_Where & " AND " & GenerateSearchFieldSQLFragment(txtArray(counter))
					End Select
				Next	'counter
			End If
		ElseIf searchParamType = "ANY" Then
			txtArray = split(searchParamTxt, " ")
			upperLim = Ubound(txtArray)
			pstrSQL_Where=pstrSQL_Where & " AND("
			If Len(searchParamTxt) > 0 Then
				For counter = 0 to (upperLim-1)
					pstrSQL_Where = pstrSQL_Where & GenerateSearchFieldSQLFragment(txtArray(counter)) & " OR "
				Next
				pstrSQL_Where = pstrSQL_Where & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%')"
			End If 
			pstrSQL_Where=pstrSQL_Where & ")"
		ElseIf searchParamType = "Exact" Then
			If Len(searchParamTxt) > 0 Then pstrSQL_Where = pstrSQL_Where & " AND " & GenerateSearchFieldSQLFragment(searchParamTxt)
		End If
	End if	'Len(searchParamTxt) > 0

	If Not isDate(DateAddedStart) Then DateAddedStart = ""
	If Not isDate(DateAddedEnd) Then DateAddedEnd = ""
	
    If isDate(DateAddedEnd) Then DateAddedEnd = dateAdd("d", 1, DateAddedEnd)
    
	'modified to be SQL Server/Access generic
    If DateAddedStart <> "" And DateAddedEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " BETWEEN " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " AND " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "
	If DateAddedStart <> "" And DateAddedEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " > " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " < " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "  

	If Not isNumeric(PriceStart) Then PriceStart = ""
	If Not isNumeric(PriceEnd) Then PriceEnd = ""
	If cblnSQLDatabase Then
		If PriceStart <> "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	Else
		If PriceStart <> "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	End If
	
	If Len(Sale) > 0 Then
		If isNumeric(Sale) Then
			If CDbl(Sale) = 0 Then
				pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodSaleIsActive = 1)"
			Else
				If cblnSQLDatabase Then
					pstrSQL_Where = pstrSQL_Where & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-convert(money,sfProducts.prodSalePrice)/convert(money,sfProducts.prodPrice)))>" & Sale & "))"
				Else
					pstrSQL_Where = pstrSQL_Where & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-CDbl([prodSalePrice])/CDbl([prodPrice])))>" & Sale & "))"
				End If
			End If
		End If
	End If

	'debugprint "pstrSQL_Where", pstrSQL_Where
	'debugprint "pstrSQL_Where", pstrSQL & pstrSQL_Where & pstrSQL_GroupBy & SearchResults_OrderBy(False)
	'Response.Flush
	
	Set rsSearch = CreateObject("ADODB.RecordSet")
	With rsSearch
		.CursorLocation = adUseClient
		.CacheSize = iVarPageSize
		.MaxRecords = iMaxRecords
		.CursorLocation = adUseClient
		
		On Error Resume Next
		.Open pstrSQL & pstrSQL_Where & pstrSQL_GroupBy & SearchResults_OrderBy(False), cnn, adOpenStatic, adLockBatchOptimistic, adCmdText 

		If Err.number <> 0 Then	' And vDebug = 1
			Response.Write "<fieldset><legend>Error in getProductIDsMatchingSearch</legend>"
			Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />"
			Response.Write "SQL: " & pstrSQL & pstrSQL_Where & pstrSQL_GroupBy & SearchResults_OrderBy(False) & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
			Err.Clear
		End If

		.PageSize = iVarPageSize
		If iPage > rsSearch.PageCount Then iPage = rsSearch.PageCount
		
		If Not .EOF Then
			If Len(iPage) > 0 Then .AbsolutePage = iPage
			mlngNumProductsFound = .RecordCount
			mstrProductIDList = .getString(adClipString, iVarPageSize, "-", "', '", "")
			mstrProductIDList = "'" & Left(mstrProductIDList, Len(mstrProductIDList) - 3)	'remove the trailing comma
			getProductIDsMatchingSearch = mstrProductIDList
		Else
			mlngNumProductsFound = 0
		End If
	End With
	Call closeObj(pobjRS)

	If vDebug = 1 Then
		Response.Write "<fieldset><legend>getProductIDsMatchingSearch</legend>"
		Response.Write "pstrSQL: " & pstrSQL & "<br /><hr>"
		Response.Write "pstrSQL_Where: " & pstrSQL_Where & "<br /><hr>"
		Response.Write "pstrSQL_GroupBy: " & pstrSQL_GroupBy & "<br /><hr>"
		Response.Write "SQL out: " & pstrSQL & pstrSQL_Where & pstrSQL_GroupBy & "<br />"
		Response.Write "mstrProductIDList: " & mstrProductIDList & "<br />"
		Response.Write "</fieldset>"
	End If
	
End Function	'getProductIDsMatchingSearch

'**********************************************************

Function BuildCustomSearchFilter(byRef strSQL, byVal searchParamType, byVal searchParamTxt, byVal searchParamCat, byVal searchParamMan, byVal searchParamVen, byVal DateAddedStart, byVal DateAddedEnd, byVal PriceStart, byVal PriceEnd, byVal Sale, byVal subCatID, byVal Ilevel)

Dim counter
Dim pstrSQL
Dim pstrSQL_GroupBy
Dim pstrSQL_Where
Dim txtArray
Dim upperLim

	Call getProductIDsMatchingSearch(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale, subCatID, Ilevel)

	If Len(mstrProductIDList) > 0 Then
		'pstrSQL_Where = categorySearchWhere(subCatID, Ilevel)
		If Len(pstrSQL_Where) > 0 Then
			strSQL = "SELECT Distinct sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, sfVendors.vendName, sfManufacturers.mfgName, sortCatDetail" _
				   & " FROM (sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId) LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID" _
				   & " WHERE sfProducts.prodID IN (" & mstrProductIDList & ")" & " And " & pstrSQL_Where
		Else
			strSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, sfVendors.vendName, sfManufacturers.mfgName" _
				& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
				& " WHERE sfProducts.prodID IN (" & mstrProductIDList & ")"
		End If
	Else
		strSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, sfVendors.vendName, sfManufacturers.mfgName" _
			& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
			& " WHERE (sfProducts.prodEnabledIsActive=1) AND (sfProducts.prodEnabledIsActive=0)"
			
		strSQL = ""
	End If
	
	'debugprint "BuildCustomSearchFilter", strSQL
	Exit Function
	
	If Application("AppName") = "StoreFrontAE" Then
	
		If True Then
			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfVendors.vendName, sfManufacturers.mfgName" _
					& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId"
			pstrSQL_Where = " WHERE (sfProducts.prodEnabledIsActive=1)"
			pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfVendors.vendName, sfManufacturers.mfgName, sfProducts.prodNamePlural, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice" & SearchResults_OrderBy(True)

		Else
			pstrSQL = "SELECT sfProducts.prodID" _
					& " FROM sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID"

			pstrSQL_Where = " WHERE (sfProducts.prodEnabledIsActive=1)"
			pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodShortDescription, sfProducts.prodNamePlural" & SearchResults_OrderBy(True)
		End If
		
		mblnSubCategorySearch = CBool(Ilevel > 1)
		If CBool(Ilevel < 2) And  Len(subCatID) > 0 And isNumeric(subCatID) Then
			pstrSQL_Where = pstrSQL_Where & " AND (sfSubCatDetail.subcatCategoryId In (Select subcatID From sfSub_Categories Where subcatCategoryId= " & subCatID & "))"
		ElseIf Len(subCatID) > 0 And isNumeric(subCatID) Then
			pstrSQL_Where = pstrSQL_Where & " " _
						  & "AND " _
						  & "(" _
						  & "  (sfSubCatDetail.subcatCategoryId= " & subCatID & ")" _
						  & " OR (sfSubCatDetail.subcatCategoryId In (Select subcatID From sfSub_Categories Where left(CatHierarchy," & Len(CStr(subCatID)) + 1 & ")='" & subCatID & "-'))" _
						  & ")"
		End If
	Else
		pstrSQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription, sfManufacturers.mfgName" _
				& " FROM sfManufacturers RIGHT JOIN (sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerID"
		If iSubCat = "ALL" Then
			pstrSQL_Where = " WHERE "
		Else
			If Len(searchParamCat) > 0 And isNumeric(searchParamCat) Then
				pstrSQL_Where = " WHERE sfProducts.prodEnabledIsActive=1 And sfProducts.prodCategoryId=" & searchParamCat & " AND "
			Else
				pstrSQL_Where = " WHERE sfProducts.prodEnabledIsActive=1"
			End If
		End If 
	End If 

    If Len(searchParamMan) > 0 And isNumeric(searchParamMan) Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts.prodManufacturerId=" & wrapSQLValue(searchParamMan, False, enDatatype_number)
    If Len(searchParamVen) > 0 And isNumeric(searchParamVen) Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts.prodVendorId=" & wrapSQLValue(searchParamVen, False, enDatatype_number)

	If Len(searchParamTxt) > 0 Then 
		If searchParamType = "ALL" Then 
			txtArray = split(searchParamTxt, " ")
			
			If searchParamTxt <> "" Then
				For counter= 0 to Ubound(txtArray)
					Select Case Request.QueryString("narrowSearch")
						Case "prodID"
							pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodID LIKE '%" & txtArray(counter) & "%') "
						Case "mfgID"
							pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodName LIKE '%" & txtArray(counter) & "%') "
						Case Else
							pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') "
					End Select
				Next	'counter
			End If
		ElseIf searchParamType = "ANY" Then
			txtArray = split(searchParamTxt, " ")
			upperLim = Ubound(txtArray)
			pstrSQL_Where=pstrSQL_Where & " AND("
			If searchParamTxt <> "" Then
				For counter=0 to (upperLim-1)
					pstrSQL_Where = pstrSQL_Where & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') OR "
				Next
				pstrSQL_Where = pstrSQL_Where & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%')"
			End If 
			pstrSQL_Where=pstrSQL_Where & ")"
		Elseif searchParamType = "Exact" Then
			If searchParamTxt <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodName LIKE '%" & searchParamTxt & "%' OR sfProducts.prodShortDescription LIKE '%" & searchParamTxt & "%' OR sfProducts.prodID ='" & searchParamTxt & "' OR sfProducts.prodDescription LIKE '%" & searchParamTxt & "%') "
		End If
	End if	'Len(searchParamTxt) > 0

	If Not isDate(DateAddedStart) Then DateAddedStart = ""
	If Not isDate(DateAddedEnd) Then DateAddedEnd = ""
	
    If isDate(DateAddedEnd) Then DateAddedEnd = dateAdd("d", 1, DateAddedEnd)
    
	'modified to be SQL Server/Access generic
    If DateAddedStart <> "" And DateAddedEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " BETWEEN " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " AND " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "
	If DateAddedStart <> "" And DateAddedEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " > " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND sfProducts." & cstrNewProductField & " < " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "  

	If Not isNumeric(PriceStart) Then PriceStart = ""
	If Not isNumeric(PriceEnd) Then PriceEnd = ""
	If cblnSQLDatabase Then
		If PriceStart <> "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (convert(money,sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	Else
		If PriceStart <> "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then pstrSQL_Where = pstrSQL_Where & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	End If
	
	If Len(Sale) > 0 Then
		If isNumeric(Sale) Then
			If CDbl(Sale) = 0 Then
				pstrSQL_Where = pstrSQL_Where & " AND (sfProducts.prodSaleIsActive = 1)"
			Else
				If cblnSQLDatabase Then
					pstrSQL_Where = pstrSQL_Where & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-convert(money,sfProducts.prodSalePrice)/convert(money,sfProducts.prodPrice)))>" & Sale & "))"
				Else
					pstrSQL_Where = pstrSQL_Where & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-CDbl([prodSalePrice])/CDbl([prodPrice])))>" & Sale & "))"
				End If
			End If
		End If
	End If
	
	If vDebug = 1 Then
		Response.Write "<fieldset><legend>BuildCustomSearchFilter</legend>"
		Response.Write "pstrSQL: " & pstrSQL & "<br /><hr>"
		Response.Write "pstrSQL_Where: " & pstrSQL_Where & "<br /><hr>"
		Response.Write "pstrSQL_GroupBy: " & pstrSQL_GroupBy & "<br /><hr>"
		Response.Write "SQL out: " & pstrSQL & pstrSQL_Where & pstrSQL_GroupBy & "<br />"
		Response.Write "</fieldset>"
	End If
	
	strSQL = pstrSQL & pstrSQL_Where & pstrSQL_GroupBy

End Function	'BuildCustomSearchFilter

'**************************************************************************************************

Sub LoadSearchParameters

	'Category information is split across two tables
	'For subcategories the following rules apply
	'txtsearchParamCat is ALWAYS the top level category
	'iLevel is always the depth; it may be empty
	'sALLSUB is a flag; either "ALL" or irrelevant
	
	txtsearchParamCat	= trim(Request.QueryString("txtsearchParamCat"))
    iLevel  = LoadRequestValue("iLevel")
	If Len(iLevel) = 0 Or iLevel = "1"  Or Not isNumeric(iLevel) then
		sALLSUB = "ALL"
		sSubCat = txtsearchParamCat
		iLevel = 1
	Else
	    sALLSUB = ""
		sSubCat = LoadRequestValue("subcat")
		iLevel = CLng(iLevel)
    End If

   ' Requests the variables depending on how the page is entered
	txtFromSearch = Trim(LoadRequestValue("txtFromSearch"))
	If txtFromSearch = "fromSearch" Then
		if iLevel = "2" and sALLSUB = "ALL" then
			txtsearchParamCat	= sSubCat
			iLevel = 1 
		end if
	Else
		if iLevel = "2" and sALLSUB = "ALL" then
			txtsearchParamCat	= sSubCat
			iLevel = 1 
		end if
	End If 

	txtsearchParamTxt	= trim(Replace(Replace(LoadRequestValue("txtsearchParamTxt"), "'", "''"), "*", ""))
	txtsearchParamType	= trim(LoadRequestValue("txtsearchParamType"))
	txtsearchParamMan	= trim(LoadRequestValue("txtsearchParamMan"))
	txtsearchParamVen	= trim(LoadRequestValue("txtsearchParamVen"))
	txtDateAddedStart	= MakeUSDate(trim(LoadRequestValue("txtDateAddedStart")))
	txtDateAddedEnd 	= MakeUSDate(trim(LoadRequestValue("txtDateAddedEnd")))
	txtPriceStart		= trim(LoadRequestValue("txtPriceStart"))
	txtPriceEnd 		= trim(LoadRequestValue("txtPriceEnd"))
	txtSale				= trim(LoadRequestValue("txtSale"))
	txtsearchParamCat	= trim(Request.QueryString("txtsearchParamCat"))
	
	If txtFromSearch = "fromSearch" Then
		if Ilevel = 2 and sALLSUB = "ALL" then
			txtsearchParamCat	= sSubCat
			Ilevel = 1 
		end if
	Else
		if Ilevel = 2 and sALLSUB = "ALL" then
			txtsearchParamCat	= sSubCat
			Ilevel = 1 
		End If
	End If 

	'Protect against invalid entries
	If Len(txtsearchParamType) = 0 Then txtsearchParamType = "ALL"
	If Len(txtsearchParamMan) = 0 Or Not isNumeric(txtsearchParamMan) Then txtsearchParamMan = "ALL"
	If Len(txtsearchParamVen) = 0 Or Not isNumeric(txtsearchParamVen) Then txtsearchParamVen = "ALL"
	If Len(txtsearchParamCat) = 0 Or Not isNumeric(txtsearchParamCat) Then txtsearchParamCat = "ALL"
	
	If Not isNumeric(txtPriceStart) Then txtPriceStart = ""
	If Not isNumeric(txtPriceEnd) Then txtPriceEnd = ""

	If Not isDate(txtDateAddedStart) Then txtDateAddedStart = ""
	If Not isDate(txtDateAddedEnd) Then txtDateAddedEnd = ""
	
	mblnCategorySearch = CBool(txtsearchParamCat <> "ALL")
	mblnManufacturerSearch = CBool(txtsearchParamMan <> "ALL")
	mblnVendorSearch = CBool(txtsearchParamVen <> "ALL")

	If Ilevel = 1 And Len(sSubCat) = 0 Then sSubCat = "ALL"

	'Determine if this page is from a search or not
	'Available types: category page, manufacturer page, vendor page, sale page, new products page, quick search/advancedSearch
	If mblnCategorySearch Then
		mbytPageType = enPageType_Category
	ElseIf mblnManufacturerSearch Then
		mbytPageType = enPageType_Mfg
	ElseIf mblnVendorSearch Then
		mbytPageType = enPageType_Vend
	ElseIf isNumeric(txtSale) And Len(txtSale) > 0 Then
		mbytPageType = enPageType_Sales
	ElseIf isDate(txtDateAddedStart) And Len(txtDateAddedStart) > 0 Then
		mbytPageType = enPageType_NewProduct
	Else
		mbytPageType = enPageType_Search
	End If

	If vDebug Then Call WriteSearchParameters	' Or True
	
End Sub	'LoadSearchParameters

'***********************************************************************************************

Sub WriteSearchParameters()

	Response.Write "<fieldset><legend>Search Parameters</legend>"
	Response.Write "txtFromSearch: " & txtFromSearch & "<br />"
	Response.Write "txtsearchParamTxt: " & txtsearchParamTxt & "<br />"
	Response.Write "txtsearchParamType: " & txtsearchParamType & "<br />"
	Response.Write "txtsearchParamMan: " & txtsearchParamMan & "<br />"
	Response.Write "txtsearchParamVen: " & txtsearchParamVen & "<br />"
	Response.Write "txtDateAddedStart: " & txtDateAddedStart & "<br />"
	Response.Write "txtDateAddedEnd: " & txtDateAddedEnd & "<br />"
	Response.Write "txtPriceStart: " & txtPriceStart & "<br />"
	Response.Write "txtPriceEnd: " & txtPriceEnd & "<br />"
	Response.Write "txtSale: " & txtSale & "<br />"
	Response.Write "txtsearchParamCat: " & txtsearchParamCat & "<br />"
	Response.Write "txtPriceEnd: " & txtPriceEnd & "<br />"
	
	Response.Write "sSubCat: " & sSubCat & "<br />"
	Response.Write "sALLSUB: " & sALLSUB & "<br />"
	Response.Write "ilevel: " & ilevel & "<br />"
	Response.Write "txtCatName: " & txtCatName & "<br />"
	Response.Write "</fieldset>"
	
End Sub	'WriteSearchParameters

'***********************************************************************************************

Sub LoadSort(byRef strSQL, byRef lngPageSize)

Dim plngPageSize

	If Len(LoadRequestValue("customSearch")) > 0 Then
		mstrHideImages = LoadRequestValue("chkHideImages")
		Session("HideImages") = mstrHideImages
	
		mbytSortBy = LoadRequestValue("sortBy")
		Session("sortBy") = mbytSortBy
		
		mbytDisplayStyle = LoadRequestValue("displayStyle")
		Session("displayStyle") = mbytDisplayStyle
	Else
		mstrHideImages = Session("HideImages")
		mbytSortBy = Session("sortBy")
		mbytDisplayStyle = Session("displayStyle")
	End If

	'Used to hardcode display style
	If Len(cstrSearchResultsDisplayTypeFixed) > 0 Then	mbytDisplayStyle = CInt(cstrSearchResultsDisplayTypeFixed)
	
	'Validate data and correct if necessary
	If Not isNumeric(mbytSortBy) And Len(mbytSortBy) > 0 Then mbytSortBy = 0	' Or Len(mbytSortBy) = 0 - length check removed to support merchant defined product ordering
	If Not isNumeric(mbytDisplayStyle) Or Len(mbytDisplayStyle) = 0 Then mbytDisplayStyle = cbytSearchResultsDisplayType
	plngPageSize = LoadRequestValue("PageSize")
	If isNumeric(plngPageSize) And Len(plngPageSize) > 0 Then
		Session("PageSize") = plngPageSize
		lngPageSize = plngPageSize
	Else
		If Len(Session("PageSize")) > 0 Then
			lngPageSize = Session("PageSize")
		Else
			lngPageSize = cssDefaultPageSize
		End If
	End If
	
	mstrSortOrder = LoadRequestValue("sortOrder")
	If Len(mstrSortOrder) = 0 Then
		If Len(Session("sortOrder")) = 0 Then
			mstrSortOrder = "ASC"
		Else
			mstrSortOrder = Session("sortOrder")
		End If
	End If
	Session("sortOrder") = mstrSortOrder

	strSQL = strSQL & SearchResults_OrderBy(False)
	
End Sub	'LoadSort

'************************************************************************************************************************************

Function SearchResults_OrderBy(byVal blnGroupBy)

Dim pstrOrderBy
Dim pstrGroupBy

	Select Case CStr(mbytSortBy)
		Case "0":	'prodID
			pstrOrderBy = " Order By sfProducts.prodID " & mstrSortOrder
		Case "1":	'prodName
			pstrOrderBy = " Order By sfProducts.prodName " & mstrSortOrder
		Case "2":	'price
			If cblnSQLDatabase then
				pstrOrderBy = " Order By convert(money,prodPrice)" & mstrSortOrder
			Else
				pstrOrderBy = " Order By cCur(prodPrice)" & mstrSortOrder
			End If
		Case Else	'use merchant defined
			pstrGroupBy = ", " & ProductSortField(DetermineSortType)
			pstrOrderBy = " Order By " & ProductSortField(DetermineSortType)
	End Select
	
	If blnGroupBy Then
		SearchResults_OrderBy = pstrGroupBy
	Else
		SearchResults_OrderBy = pstrOrderBy
	End If

End Function	'SearchResults_OrderBy

'****************************************************************************************************************

Function DetermineSortType()
'0 - Category
'1 - Manufacturer
'2 - Vendor
'3 - Category from sfSubCatDetail

	If mblnCategorySearch Then
		If cblnSF5AE Then
			DetermineSortType = 3
			DetermineSortType = 0	'revert to SE method; AE method results in duplicates
		Else
			DetermineSortType = 0
		End If
	ElseIf mblnManufacturerSearch Then
		DetermineSortType = 1
	ElseIf mblnVendorSearch Then
		DetermineSortType = 2
	Else
		DetermineSortType = 0
	End If
	
End Function	'DetermineSortType

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
%>