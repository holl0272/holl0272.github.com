<!-- #include file = "incsearchresult_custom.asp" -->
<%
'********************************************************************************
'*   Customized file for Sandshot Software Master Template		                *
'*   Release Version:	1.0.0	                                                *
'*   Release Date:		February 5, 2003			                            *
'*   Revision Date:		February 5, 2003			                            *
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*                                                                              *
'*   This file contains code copyrighted by Sandshot Software and				*
'*   LaGarde Incorporated. It has been modified by Sandshot Software which		*
'*   retains all rights to the affected code. The base code remains the property*
'*   of LaGarde Incorporated and is redistributed here with permission.			*
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.3

'@FILENAME: incSearCHrESULTS.asp
	 
'Access Version

'@DESCRIPTION:   functions to return search results

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

'Modified 11/20/01 
'Storefront Ref#'s: 131 'JF

'***********************************************************************************************

Function getManufacturersList(byVal iValue)

Dim i
Dim plngNumItems
Dim pstrOut

	If loadManufacturerArray Then
		plngNumItems = UBound(maryManufacturers)
		For i = 0 To plngNumItems
			If Not maryManufacturers(i)(8) Then pstrOut = pstrOut & "<OPTION value=" & maryManufacturers(i)(0) & " " & isSelected(maryManufacturers(i)(0) = iValue) & ">" & maryManufacturers(i)(1) & "</OPTION>"
		Next 'i
	Else
		'pstrOut = "<option value="""">No Manufacturers</option>"
	End If

	getManufacturersList = pstrOut
	
End Function	'getManufacturersList

'***********************************************************************************************

Function getVendorList(byVal iValue)

Dim i
Dim plngNumItems
Dim pstrOut

	If loadVendorArray Then
		plngNumItems = UBound(maryVendors)
		For i = 0 To plngNumItems
			If Not maryVendors(i)(8) Then pstrOut = pstrOut & "<OPTION value=" & maryVendors(i)(0) & " " & isSelected(maryVendors(i)(0) = iValue) & ">" & maryVendors(i)(1) & "</OPTION>"
		Next 'i
	Else
		'pstrOut = "<option value="""">No Manufacturers</option>"
	End If

	getVendorList = pstrOut
	
End Function	'getVendorList

'***********************************************************************************************

Function getCategoryList(byVal iValue)

Dim pobjRS
Dim pstrSQL
Dim sList
Dim intId
Dim i
	
	pstrSQL = "SELECT DISTINCT catID, catName FROM sfCategories Order By CatName"
	Set pobjRS = GetRS(pstrSQL)
	sList = ""
	For i = 1 To pobjRS.RecordCount
		intId = CStr(Trim(pobjRS.Fields(0).Value & ""))
		If CStr(iValue) = intId Then
			slist = slist & "<OPTION value=" & intId & " selected>" & pobjRS.Fields(1).Value & "</OPTION>"
		Else
			slist = slist & "<OPTION value=" & intId & ">" & pobjRS.Fields(1).Value & "</OPTION>"
		End If 
		pobjRS.MoveNext
	Next 'i				

	Call closeObj(pobjRS)

	getCategoryList = sList
	
End Function	'getCategoryList


'***********************************************************************************************

'--------------------------------------------------------------------------------------------------

Function getSubCategoryList(ilevel,subcatID)

Dim rsSubCategoryList
Dim sList,iLen
dim sSQl,sHierarchy
dim MainCatID

	if instr(subcatid, "-") > 0 Or Len(iLevel) = 0 Then
		getSubCategoryList = ""
		exit function
	end if  
	
	if ilevel = 1 And subCatID <> "ALL" then 
		MainCatID = setSubcatId(subCatId,"subcatCategoryId","subcatCategoryId")
	elseif ilevel > 1 And subCatID <> "ALL" then
		MainCatID = setSubcatId(subCatId,"subCatId","subcatCategoryId")
	end if  

	if  subCatID <> "ALL" then 
		sHierarchy = getCatHierarchy(subcatID)
		iLen = len(sHierarchy)
		if Ilevel = 1 then
			sSQl = "SELECT Distinct subcatCategoryId, subcatID ,SubcatName,Bottom  FROM sfSub_Categories Where Depth = " & iLevel & " And subcatCategoryId = " & MainCatID
		else
			sSQl = "SELECT Distinct subcatCategoryId, subcatID ,SubcatName,Bottom  FROM sfSub_Categories" _
				 & " Where Depth = " & iLevel & " And subcatCategoryId = " & MainCatID &	" AND LEFT(CatHierarchy," & iLen & ") = '" & sHierarchy & "'"
		end if
		sSQl = sSQl & " Order By SubcatName"

		Set rsSubCategoryList = CreateObject("ADODB.RecordSet")
		rsSubCategoryList.Open sSQL , cnn, adOpenForwardOnly, adLockOptimistic, adCmdText

		if rsSubCategoryList.EOF = false and rsSubCategoryList.BOF =false then
			Do While Not rsSubCategoryList.EOF
				if rsSubCategoryList.Fields("Bottom") = 1 then
					sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("subcatID") & "-bottom>" & rsSubCategoryList.Fields("subcatName") & "</OPTION>"
				else    
					sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("subcatID") & ">" & rsSubCategoryList.Fields("subcatName") & "</OPTION>"
				end if 	
				rsSubCategoryList.MoveNext                                   
			Loop
		else
			sList = ""
		end if 
	else
		sSQl = "SELECT DISTINCT catID, catName  FROM sfCategories " 
		Set rsSubCategoryList = CreateObject("ADODB.RecordSet")
		rsSubCategoryList.Open sSQL , cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if rsSubCategoryList.EOF = false and rsSubCategoryList.BOF =false then
			Do While Not rsSubCategoryList.EOF
				sList = sList & "<OPTION value=" & rsSubCategoryList.Fields("catID") & ">" & rsSubCategoryList.Fields("catName") & "</OPTION>"
				rsSubCategoryList.MoveNext                                   
			Loop
		else
			sList = ""
		end if   
	end if 

	rsSubCategoryList.Close
	Set rsSubCategoryList = nothing 
	
	getSubCategoryList = sList
	
End Function	'getSubCategoryList

'--------------------------------------------------------------------------------------------------

Function GenerateSearchFieldSQLFragment(byVal strSearchText)

Dim SearchFieldCounter
Dim SQL_SearchField
Dim parySearchFields

	If cblnSearchAttributes Then
		parySearchFields = Array("sfProducts.prodName", "sfProducts.prodShortDescription", "sfProducts.prodID", "sfProducts.prodDescription", "sfAttributes.attrName", "sfAttributeDetail.attrdtName")
	Else
		parySearchFields = Array("sfProducts.prodName", "sfProducts.prodShortDescription", "sfProducts.prodID", "sfProducts.prodDescription")
	End If

	SQL_SearchField = ""
	For SearchFieldCounter = 0 To UBound(parySearchFields)
		If Len(SQL_SearchField) = 0 Then
			SQL_SearchField = "(" & "(" & parySearchFields(SearchFieldCounter) & " LIKE '%" & strSearchText & "%')"
		Else
			SQL_SearchField = SQL_SearchField & " OR " & "(" & parySearchFields(SearchFieldCounter) & " LIKE '%" & strSearchText & "%')"
		End If
	Next
	SQL_SearchField = SQL_SearchField & ")"
	
	GenerateSearchFieldSQLFragment = SQL_SearchField

End Function	'GenerateSearchFieldSQLFragment

'--------------------------------------------------------------------------------------------------

Function getProductSQLAE(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale, subCatID, Ilevel)

Dim upperLim, SQL, counter, txtArray
Dim pstrSQL_GroupBy

	if instr(subcatid, "bottom") > 0 then
		subcatid = left(subcatID,instr(subcatId,"-")-1)
	end if
	searchParamTxt = Replace(searchParamTxt, "*", "")
	sSubcat =  subCatID 

	if iLevel = 1 and subCatID <> "ALL"   then 
		subCatId = setSubcatId(subCatId,"subcatCategoryId","subcatCategoryId")
	end if

	if iLevel = 1 then sALLSUB = "ALL"
	if Len(subCatID) = 0 Then subCatID = "ALL"

	if subCatID = "ALL" Then
		SQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice" _
			& " FROM ((sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
			& " WHERE "

		'need the group by to eliminate duplicates
		pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum"
		'debugprint "ALL", SQL
	ElseIf instr(subCatID,";") > 0  Then
		'this section may never be pulled due to revised template
		SQL = " SELECT sfProducts.ProdID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice,sfSub_Categories.CatHierarchy," _
			& "        sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription" _
			& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN ((sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
			& " WHERE (sfSubCatDetail.subcatCategoryId IN (Select subcatCategoryId From sfSubCatDetail Where subcatCategoryId = " & GetSubCatIDs(sSubCat) & ")) AND " 

		'need the group by to eliminate duplicates
		pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodNamePlural, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, prodPLPrice, prodPLSalePrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfSub_Categories.CatHierarchy, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, subcatDescription As catDescription, subcatImage As catImage, sfVendors.vendName, sfManufacturers.mfgName, subcatDescription, subcatImage"
		'debugprint "subCatID", SQL
	Else
		SQL = " SELECT sfProducts.ProdID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice,sfSub_Categories.CatHierarchy," _
			& "        sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, subcatDescription As catDescription, subcatImage As catImage, sfVendors.vendName, sfManufacturers.mfgName" _
			& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN ((sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId"

		if sALLSUB <> "ALL" Then
			SQL = SQL & " WHERE (sfSubCatDetail.subcatCategoryId IN (Select subcatCategoryId From sfSubCatDetail Where subcatCategoryId = " & GetSubCatIDs(sSubCat) & ")) AND " 
		else
			SQL = SQL & " WHERE sfSubCatDetail.subcatCategoryId IN (Select subcatID From sfSub_Categories Where subcatCategoryId= " & sSubCat & ") AND " 
		end if
		
		'need the group by to eliminate duplicates
		pstrSQL_GroupBy = " GROUP BY sfProducts.prodID, sfProducts.prodName, sfProducts.prodNamePlural, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, prodPLPrice, prodPLSalePrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfSub_Categories.CatHierarchy, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, subcatDescription, subcatImage, sfVendors.vendName, sfManufacturers.mfgName"
		'debugprint "Else", SQL
		mblnSubCategorySearch = True
	End If 

	if Len(searchParamTxt) > 0 Then 
		If searchParamType = "ALL" Then 
			txtArray = split(searchParamTxt, " ")
			
			If searchParamTxt <> "" Then
				For counter=0 to Ubound(txtArray)
					Select Case Request.QueryString("narrowSearch")
						Case "prodID"
							SQL = SQL & " (sfProducts.prodID LIKE '%" & txtArray(counter) & "%') AND "
						Case "mfgID"
							SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%') AND "
						Case Else
							SQL = SQL & GenerateSearchFieldSQLFragment(txtArray(counter)) & " AND "
							'SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') AND "
					End Select
				Next	'counter
			End If
				 
		'	Response.Write " A-1  <br />"	
		Elseif searchParamType = "ANY" Then
			'Response.Write " B  <br />"
			txtArray = split(searchParamTxt, " ")
			upperLim = Ubound(txtArray)
			SQL=SQL & "("
			If searchParamTxt <> "" Then
				For counter=0 to (upperLim-1)
					SQL = SQL & GenerateSearchFieldSQLFragment(txtArray(counter)) & " OR "
					'SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%') OR "
				Next
				SQL = SQL & " (sfProducts.prodName LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodShortDescription LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodID LIKE '%" & txtArray(counter) & "%' OR sfProducts.prodDescription LIKE '%" & txtArray(counter) & "%')"
			End If 
			SQL=SQL & ")"
		'Response.Write " B-1  <br />"	
			SQL = SQL & " And "		
		Elseif searchParamType = "Exact" Then
		' Response.Write " C  <br />"
		'if searchParamTxt <> "" Then SQL = SQL & " (sfProducts.prodName LIKE '%" & searchParamTxt & "%' OR sfProducts.prodShortDescription LIKE '%" & searchParamTxt & "%' OR sfProducts.prodID ='" & searchParamTxt & "' OR sfProducts.prodDescription LIKE '%" & searchParamTxt & "%') AND "
		if searchParamTxt <> "" Then SQL = SQL & GenerateSearchFieldSQLFragment(searchParamTxt) & " AND "
		Else 
		'Response.Write " D  <br />"
			SQL = SQL & " WHERE "
		End If
	end if	'Len(searchParamTxt) > 0

	If searchParamMan = "ALL" Then 
		SQL = SQL &  " sfProducts.prodManufacturerId > 0"
	ElseIf (isNumeric(searchParamMan) And Len(searchParamMan) > 0) Then
		SQL = SQL & "  sfProducts.prodManufacturerId = " & searchParamMan
	End If 
	If searchParamVen = "ALL" Then 
		SQL = SQL & " AND sfProducts.prodVendorId > 0"
	ElseIf (isNumeric(searchParamVen) And Len(searchParamVen) > 0) Then
		SQL = SQL & " AND sfProducts.prodVendorId = " & searchParamVen
	End If	
	
	If Not isDate(DateAddedStart) Then DateAddedStart = ""
	If Not isDate(DateAddedEnd) Then DateAddedEnd = ""
	
    If isDate(DateAddedEnd) Then DateAddedEnd = dateAdd("d",1,DateAddedEnd)
    
	'modified to be SQL Server/Access generic
    If DateAddedStart <> "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " BETWEEN " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " AND " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "
	If DateAddedStart <> "" And DateAddedEnd = "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " > " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " < " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "  

	If Not isNumeric(PriceStart) Then PriceStart = ""
	If Not isNumeric(PriceEnd) Then PriceEnd = ""
	If cblnSQLDatabase Then
		If PriceStart <> "" And PriceEnd <> "" Then SQL = SQL & " AND (convert(money,sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then SQL = SQL & " AND (convert(money,sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then SQL = SQL & " AND (convert(money,sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and convert(money,sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	Else
		If PriceStart <> "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
		If PriceStart <> "" And PriceEnd = "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
		If PriceStart = "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 
	End If
	
	If Len(Sale) > 0 Then
		If isNumeric(Sale) Then
			If CDbl(Sale) = 0 Then
				SQL = SQL & " AND sfProducts.prodSaleIsActive = 1 "
			Else
				If cblnSQLDatabase Then
					SQL = SQL & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-convert(money,sfProducts.prodSalePrice)/convert(money,sfProducts.prodPrice)))>" & Sale & ")) "
				Else
					SQL = SQL & " AND ((sfProducts.prodSaleIsActive=1) AND ((100*(1-CDbl([prodSalePrice])/CDbl([prodPrice])))>" & Sale & ")) "
				End If
				'Response.Write "SQL: " & SQL & "<br />"
			End If
		End If
	End If

	SQL = SQL & " AND sfProducts.prodEnabledIsActive = 1 "
	
	If Len(pstrSQL_GroupBy) > 0 Then SQL = SQL & pstrSQL_GroupBy
	
	getProductSQLAE = SQL
	
	'debugprint "SQL", SQL

End Function	'getProductSQLAE

'--------------------------------------------------------------------------------------------------

Function getProductSQL(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale)

Dim upperLim, SQL, counter, txtArray

	searchParamTxt = Replace(searchParamTxt, "*", "") 
	SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
		& "FROM sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID WHERE "

	SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription, sfManufacturers.mfgName " _
		& "FROM sfManufacturers RIGHT JOIN (sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID) ON sfManufacturers.mfgID = sfProducts.prodManufacturerID WHERE "
	'create where statement
	If searchParamType = "ALL" Then
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		If searchParamTxt <> "" Then
			For counter=0 to (upperLim-1)
				SQL = SQL &  " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') AND "
			Next
			SQL = SQL & " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') " 
		End If
		 	
	Elseif searchParamType = "ANY" Then
		txtArray = split(searchParamTxt, " ")
		upperLim = Ubound(txtArray)
		
		If searchParamTxt <> "" Then
		    SQL=SQL & "("  '#487 
     		For counter=0 to (upperLim-1)
				SQL = SQL &  " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%') OR "
			Next
			SQL = SQL & " (prodName LIKE '%" & txtArray(counter) & "%' OR prodShortDescription LIKE '%" & txtArray(counter) & "%' OR prodID LIKE '%" & txtArray(counter) & "%' OR prodDescription LIKE '%" & txtArray(counter) & "%')"
	    	SQL=SQL & ")"
		End If 
	
	Elseif searchParamType = "Exact" Then
			If searchParamTxt <> "" Then SQL = SQL & " (prodName LIKE '%" & searchParamTxt & "%' OR prodShortDescription LIKE '%" & searchParamTxt & "%' OR prodID ='" & searchParamTxt & "' OR prodDescription LIKE '%" & searchParamTxt & "%') "
	End If
	
	If searchParamCat = "ALL" Then
		If searchParamTxt = "" Then
			SQL = SQL & " prodCategoryId > 0"
		Else 
			SQL = SQL & " AND prodCategoryId > 0"
		End If 
	Else
		If searchParamTxt = "" Then 
			SQL = SQL & " prodCategoryId = " & searchParamCat
		Else
			SQL = SQL & " AND prodCategoryId = " & searchParamCat
		End If 
	End If 
	If searchParamMan = "ALL" Then 
		SQL = SQL &  " AND prodManufacturerId > 0"
	Else
		SQL = SQL & " AND prodManufacturerId = " & searchParamMan
	End If 
	If searchParamVen = "ALL" Then 
		SQL = SQL & " AND prodVendorId > 0"
	Else
		SQL = SQL & " AND prodVendorId = " & searchParamVen
	End If	
	If  DateAddedEnd <> ""  then
     DateAddedEnd = dateAdd("d",1,DateAddedEnd)
    end if  
    
	
	'modified to be SQL Server/Access generic
    If DateAddedStart <> "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " BETWEEN " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " AND " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "
	If DateAddedStart <> "" And DateAddedEnd = "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " > " & wrapSQLValue(DateAddedStart, False, enDatatype_date) & " " 
	If DateAddedStart = "" And DateAddedEnd <> "" Then SQL = SQL & " AND sfProducts." & cstrNewProductField & " < " & wrapSQLValue(DateAddedEnd, False, enDatatype_date) & " "  

	'djp log 201
	If PriceStart <> "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) BETWEEN " & CDbl(PriceStart) & " AND " & CDbl(PriceEnd) & ")) " 
	If PriceStart <> "" And PriceEnd = "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) >= " & CDbl(PriceStart) & " or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) >= " & CDbl(PriceStart) & ")) " 
	If PriceStart = "" And PriceEnd <> "" Then SQL = SQL & " AND (CDbl(sfProducts.prodPrice) <= " & CDbl(PriceEnd) & "  or (sfProducts.prodSaleIsActive=1 and CDbl(sfProducts.prodSalePrice) <= " & CDbl(PriceEnd) & ")) " 

	
	If Sale <> "" Then SQL = SQL & " AND prodSaleIsActive = 1 "
	SQL = SQL & " AND prodEnabledIsActive = 1 "
	getProductSQL = SQL
End Function

Function getAttributeSQL(rsSearch, iPageSize, iPage)

Dim counter, rs, SQL

	If isObject(rsSearch) Then
		If Not rsSearch.EOF Then
			' Clone rsSearch so it is not manipulated by the Function
			Set rs = CreateObject("ADODB.RecordSet")
			Set rs = rsSearch.Clone 
		
			
			rs.AbsolutePosition  = rsSearch.AbsolutePosition   
			'modified for Sandshot Software's Attribute Extender
			'SQL = "SELECT attrID, attrName, attrProdID FROM sfAttributes WHERE "
			SQL = "SELECT attrID, attrName, attrProdID, attrDisplayStyle FROM sfAttributes WHERE "
			
			For counter = 1 to iPageSize
				SQL = SQL & "attrProdId = '" & rs.Fields("prodID") & "' OR "
				rs.MoveNext
			Next
			closeObj(rs)
			SQL = Mid(SQL, 1, len(SQL)-3)
			'modified for Sandshot Software's Attribute Extender
			'getAttributeSQL = SQL & " ORDER BY attrName"
			getAttributeSQL = SQL & " ORDER BY attrDisplayOrder"
		Else
			getAttributeSQL = ""
		End If
	
	Else
		If Len(rsSearch) > 0 Then
			getAttributeSQL = "SELECT attrID, attrName, attrProdID, attrDisplayStyle FROM sfAttributes WHERE attrProdID In (" & rsSearch & ")" _
							& " ORDER BY attrDisplayOrder"
		Else
			getAttributeSQL = ""
		End If 
	End If
	
End Function

Function getAttributeDetailSQL(rs)

Dim SQL

	If Not rs.EOF Then
		SQL = "SELECT attrdtID, attrdtAttributeId, attrdtName, attrdtPrice, attrdtType, attrdtOrder FROM sfAttributeDetail WHERE "
		Do While Not rs.EOF
			SQL = SQL & " attrdtAttributeId = " & rs.Fields("attrID") & " OR "
			rs.MoveNext
		Loop
		rs.MoveFirst
		SQL = Mid(SQL, 1, len(SQL)-3)
		getAttributeDetailSQL = SQL & " ORDER BY attrdtOrder"
	Else
		getAttributeDetailSQL = ""
	End If 

End Function

Function getCategorySQL(txtCategory)
	getCategorySQL = "SELECT catName FROM sfCategories WHERE catID = " & txtCategory
End Function

Function bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, sFromPage)

Dim txtPage, icounter, output, iStart, iEnd, sLink,iLoop

	output = ""
	
	if Not cblnSF5AE then
		txtPage = "&txtsearchParamTxt=" & Server.URLEncode(txtsearchParamTxt) & "&txtsearchParamType=" & txtsearchParamType & "&txtsearchParamCat=" _
			& txtsearchParamCat & "&txtsearchParamMan=" & txtsearchParamMan & "&txtsearchParamVen=" & txtsearchParamVen _
			& "&txtDateAddedStart=" & txtDateAddedStart & "&txtDateAddedEnd=" & txtDateAddedEnd _
			& "&txtPriceStart=" & txtPriceStart & "&txtPriceEnd=" & txtPriceEnd & "&txtSale=" & txtSale
	else
		txtPage =""
		  For iLoop = 1 to Request.QueryString.Count 
		    If lcase(Request.QueryString.Key(iLoop)) <> "page" then  
		     txtPage = txtPage &  "&" & Request.QueryString.Key(iLoop) & "=" & Request.QueryString.Item(iLoop)  
		    End if 
		  Next
		 txtPage = replace(txtPage," ","+")
	end if

	Select Case LCase(sFromPage)
		Case "salespage", "salespage.asp"
			sLink = "<a href=salespage.asp?PAGE="
			txtPage = ""
		Case "newproducts", "newproduct.asp"
			sLink = "<a href=newproduct.asp?PAGE="
			txtPage = ""
		Case Else
			sLink = "<a href=" & Replace(Request.ServerVariables("SCRIPT_NAME"),"/","") & "?PAGE="
	End Select
	
	If iPage <> "1" Then
	    output = output &  "<font color=black>&lt;&lt; " & sLink & iPage - 1 & txtPage &">Previous</a> | "
	Else 
	    output = output &  "<font color=silver>&lt;&lt; Previous</font> | "
	End If
	
	'Two cases, less than ten pages or more than ten pages total                
	If iNumOfPages > 10 Then 'Four cases inbeded 
		'First case, first ten pages
		If iPage <= 10 Then
			If iPage <> 10 Then
				For icounter = 1 to 9
				    If iCounter = CInt(iPage) Then
				        output = output &  iCounter & " | "
				    Else
				        output = output & sLink & icounter & txtPage & ">" & icounter & "</a> | "
				    End If                      
				Next
					output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
			Else
				If iNumOfPages < 20 Then
					For icounter = 10 to iNumOfPages
					    If iCounter = CInt(iPage) Then
					        output = output &  iCounter & " | "
					    Else
					        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
					    End If                      
					Next
				Else
					For icounter = 10 to 19
					    If iCounter = CInt(iPage) Then
					        output = output &  iCounter & " | "
					    Else
					        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
					    End If                      
					Next
						output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
				End If 
			End If 
		'rare case when the number of pages is divisable to records per page
		ElseIf iPage <= (iNumOfPages - (iNumOfPages mod 10)) AND iPage > iNumOfPages-iPageSize AND (iNumOfPages mod iPageSize) = 0 Then  
			For icounter = iNumOfPages-9 to iNumOfPages
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
		'Case for the inbetween areas ie 10-20 20-30... 
		ElseIf iPage < (iNumOfPages - (iNumOfPages mod 10)) Then
			If iPage mod 10 = 0 Then
				iStart = iPage
				iEnd = iPage + 9
			Else
				iStart = (iPage - (iPage mod 10))
				iEnd = iStart + 9
			End If  
			For icounter = iStart to iEnd
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
			output = output &  sLink & icounter & txtPage & ">" & icounter & "...</a> | "
		'Case when last few pages is less then ten
		Else
			For icounter = (iPage - (iPage mod 10)) to iNumOfPages
			    If iCounter = CInt(iPage) Then
			        output = output &  iCounter & " | "
			    Else
			        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
			    End If                      
			Next
		End If
	'If total number of pages is less than ten
	Else             
		For icounter = 1 to iNumOfPages
		    If icounter = CInt(iPage) Then
		        output = output &  iCounter & " | "
		    Else
		        output = output &  sLink & icounter & txtPage & ">" & icounter & "</a> | "
		    End If                      
		Next
	End If 
	                
	If CInt(iNumOfPages) <> CInt(iPage) Then 
	    output = output &  sLink & iPage + 1 & txtPage &">Next</a><font color=black> &gt;&gt;"
	Else
	    output = output &  "<font color=silver>Next &gt;&gt;</font>"
	End If 
	
bottomPaging = output
End Function

'--------------------------------------------------------------------------------------------------

Function GetSubCatIDs(vID)

dim rstSubCat,sSQL ,sHierarchy,iLen
dim tempID,sCriteria

	Set rstSubCat = CreateObject("ADODB.RecordSet")

	if Instr(vId,";")> 0 then
		sSql = "Select SubCatId  From sfSub_Categories Where subcatCategoryId = " & Cint(left(vID,len(instr(vid,";")-1))) _
		& " AND left(CatHierarchy,4)= '" & "none" & "'"
		rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
		 
		vID = rstSubCat("SubCatID")
		rstSubCat.Close 
	end if
	
	sSql = "Select CatHierarchy,hasprods From sfSub_Categories Where SubcatID = " & vID
	rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
	if rstSubCat.EOF =true and rstSubCat.BOF = true then
	else
		if rstSubCat("hasprods") = 1 then
			GetSubCatIDs =vID
		else 
			sHierarchy = rstSubCat("CatHierarchy")
			iLen = len(sHierarchy)
			sSql = "Select SubCatID,CatHierarchy,hasprods From sfSub_Categories Where left(CatHierarchy," & iLen & ") = '" & sHierarchy & "' AND Hasprods = 1 AND Depth > 0"
			rstSubCat.Close
			rstSubCat.Open sSql, cnn,adOpenStatic ,adLockReadOnly , adCmdText	
			if rstSubCat.EOF =true and rstSubCat.BOF = true then
			GetSubCatIDs = vID
			else
				tempID =""
				sCriteria = " OR subcatCategoryId ="
				while rstSubCat.EOF =false
					tempID = tempID & rstSubCat("SubCatID") & sCriteria  
					rstSubCat.MoveNext 
				wend     
				tempID = left(tempID,len(tempID) - len(sCriteria))
				GetSubCatIDs = tempID 
			end if 
		end if 
	end if
	if GetSubCatIDs ="" then
		GetSubCatIDs = vID
	end if

End Function	'GetSubCatIDs

'--------------------------------------------------------------------------------------------------

Function GetFullPath(Vdata,justMain,subCatID) 
Dim sSql ,X
Dim iCatId,sCriteria
Dim sFirst
Dim rst,rsCat,rsSubCat
Dim arrTemp ,bMain


If subCatID = "ALL" Then
		 sSql = "SELECT sfSubCatDetail.ProdID, sfSub_Categories.CatHierarchy" _
		 & " FROM sfSubCatDetail INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
		 & "  Where sfSubCatDetail.ProdID = '" & vData & "'"
		Set rst = CreateObject("ADODB.RecordSet")
		rst.open sSql, cnn,adOpenStatic,adLockReadOnly ,1 
        ' Response.Write ssql

		If rst.eof = false then 
		 sCriteria = rst("CatHierarchy")
		else
		 GetFullPath = "No Category"
		 rst.close
		 set rst = nothing
		 exit Function
		End if
		rst.close
		set rst = nothing
Else
 sCriteria = vData
End if

bMain = false
 if left(sCriteria,4)= "none" then
  bMain = True
  arrTemp = split(sCriteria,"-")
  sCriteria = arrtemp(1)
 elseif sCriteria = "" then
   GetFullPath = "" 
  exit function
 elseif instr(sCriteria,"-") = 0  then
    sCriteria = sCriteria 
 end if 
  arrTemp = split(sCriteria,"-")
 Set rsCat = CreateObject("ADODB.RecordSet")
 Set rsSubCat = CreateObject("ADODB.RecordSet")
  rsSubCat.Open "sfSub_Categories",cnn,adOpenStatic,adLockReadOnly ,adcmdtable 
   For X = 0 To UBound(arrTemp)
     rsSubCat.Requery
     if arrTemp(X)<> "" then
      rsSubCat.Find "SubCatId = " & CInt(arrTemp(X))
      GetFullPath = GetFullPath & rsSubCat("SubCatName") & "-"
     end if
   Next
  sSql  = "Select catName From sfCategories Where catId =" & rsSubCat("subcatCategoryId")   
 rsCat.Open sSql,cnn,adOpenStatic,adLockReadOnly ,adcmdText
 if justmain = 1 then
    GetFullPath = rsCat("catName")
 else 
   On error Resume next
   if bMain = True Then
      GetFullPath = rsCat("catName")
   else
     GetFullPath = rsCat("catName") & "-" &  Left(GetFullPath, Len(GetFullPath) - 1)
   end if 
 end if
 Set rsCat = Nothing
 Set rsSubCat = Nothing
 Exit Function
End Function

'***********************************************************************************************

function setSubcatId(iCatId, sCriteria, returnField)

dim rst,sSql

	If Len(iCatID) = 0 Or Not isNumeric(iCatId) Then Exit Function

	sSql = "Select subCatID,subcatCategoryId from sfSub_Categories where " & sCriteria & " = "  & iCatid
	Set rst = CreateObject("ADODB.RecordSet")
    rst.CursorLocation = 2 'adUseClient
	rst.Open ssql,cnn,3,1,1
	if rst.eof then
		setSubcatId = icatId
	else  
		setSubcatId = rst(returnField) 
	end if
	rst.Close
	set rst = nothing
	
end function

'***********************************************************************************************

function getCatHierarchy(vID)

dim rst,sSql

	If Len(vID) = 0 Or Not isNumeric(vID) Then Exit Function

	sSql = "Select CatHierarchy from sfSub_Categories where subcatID = "  & vID
	Set rst = CreateObject("ADODB.RecordSet")
    rst.CursorLocation = 2 'adUseClient
	rst.Open ssql,cnn,3,1,1
	if rst.eof then
		getCatHierarchy ="vID"
	else
		getCatHierarchy = rst("CatHierarchy") 
	end if
	rst.Close
	set rst = nothing
 
end function

%>



