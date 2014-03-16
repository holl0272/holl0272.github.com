<!--#include file="ssCategorySearchSE.asp"-->
<!--#include file="ssCategorySearchAE.asp"-->
<%
Dim mclsCategory:	mclsCategory = Null	'Set to null otherwise initial isObject() test returns true
Dim mstrCurrentCategory
Dim mstrCurrentSubCategory
Dim mstrCurrentSubSubCategory

'***********************************************************************************************
'***********************************************************************************************

Class sfSearch

Private pblnStoredSearch
Private pblnUseShortLink
Private p_arySearchCriteria(13)
Private enSearchCriteria_iLevel
Private enSearchCriteria_subcat
Private enSearchCriteria_txtFromSearch
Private enSearchCriteria_txtsearchParamTxt
Private enSearchCriteria_txtsearchParamType
Private enSearchCriteria_txtsearchParamMan
Private enSearchCriteria_txtsearchParamVen
Private enSearchCriteria_txtDateAddedStart
Private enSearchCriteria_txtDateAddedEnd
Private enSearchCriteria_txtPriceStart
Private enSearchCriteria_txtPriceEnd
Private enSearchCriteria_txtSale
Private enSearchCriteria_txtCatName
Private enSearchCriteria_txtsearchParamCat

	'***********************************************************************************************

	Private Sub class_Terminate()
		On Error Resume Next
	End Sub
	Private Sub class_Initialize()
	
		enSearchCriteria_iLevel = 0
		enSearchCriteria_subcat = 1
		enSearchCriteria_txtFromSearch = 2
		enSearchCriteria_txtsearchParamTxt = 3
		enSearchCriteria_txtsearchParamType = 4
		enSearchCriteria_txtsearchParamMan = 5
		enSearchCriteria_txtsearchParamVen = 6
		enSearchCriteria_txtDateAddedStart = 7
		enSearchCriteria_txtDateAddedEnd = 8
		enSearchCriteria_txtPriceStart = 9
		enSearchCriteria_txtPriceEnd = 10
		enSearchCriteria_txtSale = 11
		enSearchCriteria_txtCatName = 12
		enSearchCriteria_txtsearchParamCat = 13
		
		pblnStoredSearch = False
		pblnUseShortLink = False
		
	End Sub	'class_Initialize
	
	Public Property Let UseShortLink(byVal vntValue)
		pblnUseShortLink = vntValue
	End Property

	Public Property Let iLevel(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_iLevel) = vntValue
	End Property
	Public Property Get iLevel
		iLevel = p_arySearchCriteria(enSearchCriteria_iLevel)
	End Property

	Public Property Let subcat(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_subcat) = Replace(vntValue, "-bottom", "")
	End Property
	Public Property Get subcat
		subcat = p_arySearchCriteria(enSearchCriteria_subcat)
	End Property

	Public Property Let txtFromSearch(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtFromSearch) = vntValue
	End Property
	Public Property Get txtFromSearch
		txtFromSearch = p_arySearchCriteria(enSearchCriteria_txtFromSearch)
	End Property

	Public Property Let txtsearchParamTxt(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtsearchParamTxt) = vntValue
	End Property
	Public Property Get txtsearchParamTxt
		txtsearchParamTxt = p_arySearchCriteria(enSearchCriteria_txtsearchParamTxt)
	End Property

	Public Property Let txtsearchParamType(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtsearchParamType) = vntValue
	End Property
	Public Property Get txtsearchParamType
		txtsearchParamType = p_arySearchCriteria(enSearchCriteria_txtsearchParamType)
	End Property

	Public Property Let txtsearchParamCat(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtsearchParamCat) = vntValue
	End Property
	Public Property Get txtsearchParamCat
		txtsearchParamCat = p_arySearchCriteria(enSearchCriteria_txtsearchParamCat)
	End Property

	Public Property Let txtsearchParamMan(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtsearchParamMan) = vntValue
	End Property
	Public Property Get txtsearchParamMan
		txtsearchParamMan = p_arySearchCriteria(enSearchCriteria_txtsearchParamMan)
	End Property

	Public Property Let txtsearchParamVen(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtsearchParamVen) = vntValue
	End Property
	Public Property Get txtsearchParamVen
		txtsearchParamVen = p_arySearchCriteria(enSearchCriteria_txtsearchParamVen)
	End Property

	Public Property Let txtDateAddedStart(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtDateAddedStart) = vntValue
	End Property
	Public Property Get txtDateAddedStart
		txtDateAddedStart = p_arySearchCriteria(enSearchCriteria_txtDateAddedStart)
	End Property

	Public Property Let txtDateAddedEnd(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtDateAddedEnd) = vntValue
	End Property
	Public Property Get txtDateAddedEnd
		txtDateAddedEnd = p_arySearchCriteria(enSearchCriteria_txtDateAddedEnd)
	End Property

	Public Property Let txtPriceStart(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtPriceStart) = vntValue
	End Property
	Public Property Get txtPriceStart
		txtPriceStart = p_arySearchCriteria(enSearchCriteria_txtPriceStart)
	End Property

	Public Property Let txtPriceEnd(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtPriceEnd) = vntValue
	End Property
	Public Property Get txtPriceEnd
		txtPriceEnd = p_arySearchCriteria(enSearchCriteria_txtPriceEnd)
	End Property

	Public Property Let txtSale(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtSale) = vntValue
	End Property
	Public Property Get txtSale
		txtSale = p_arySearchCriteria(enSearchCriteria_txtSale)
	End Property

	Public Property Let txtCatName(byVal vntValue)
		p_arySearchCriteria(enSearchCriteria_txtCatName) = vntValue
	End Property
	Public Property Get txtCatName
		txtCatName = p_arySearchCriteria(enSearchCriteria_txtCatName)
	End Property

	'***********************************************************************************************

	Public Property Get HasStoredSearch
		HasStoredSearch = pblnStoredSearch
	End Property

	'***********************************************************************************************

	Public Sub SaveSearchCriteria(byVal strSessionVariableName)
		Session(strSessionVariableName) = p_arySearchCriteria
	End Sub	'SaveSearchCriteria

	'***********************************************************************************************

	Public Sub LoadSavedSearchCriteria(byVal strSessionVariableName)
	
	Dim i
	Dim paryTemp
	
		paryTemp = Session(strSessionVariableName)
		If isArray(paryTemp) Then
			pblnStoredSearch = True
			For i = 0 To UBound(paryTemp)
				If UBound(p_arySearchCriteria) >= i Then p_arySearchCriteria(i) = paryTemp(i)
			Next	'i
			
			If cblnDebugCategorySearchToolAEAddon Then WriteSearchCriteria
		End If
	End Sub	'SaveSearchCriteria

	'***********************************************************************************************

	Public Function SearchLinkParameters(byVal blnCategory)
	
	Dim pstrTemp
	
		'Check for defaults
		If Len(p_arySearchCriteria(enSearchCriteria_txtsearchParamMan)) = 0 Then p_arySearchCriteria(enSearchCriteria_txtsearchParamMan) = "ALL"
		If Len(p_arySearchCriteria(enSearchCriteria_txtsearchParamVen)) = 0 Then p_arySearchCriteria(enSearchCriteria_txtsearchParamVen) = "ALL"
		
		If pblnUseShortLink Then
			pstrTemp = ""
		Else
			pstrTemp = "txtsearchParamType=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamType) _
					& "&amp;txtsearchParamTxt=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamTxt) _
					& "&amp;txtsearchParamMan=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamMan) _
					& "&amp;txtsearchParamVen=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamVen) _
					& "&amp;txtDateAddedStart=" & p_arySearchCriteria(enSearchCriteria_txtDateAddedStart) _
					& "&amp;txtDateAddedEnd=" & p_arySearchCriteria(enSearchCriteria_txtDateAddedEnd) _
					& "&amp;txtPriceStart=" & p_arySearchCriteria(enSearchCriteria_txtPriceStart) _
					& "&amp;txtPriceEnd=" & p_arySearchCriteria(enSearchCriteria_txtPriceEnd) _
					& "&amp;txtSale=" & p_arySearchCriteria(enSearchCriteria_txtSale) _
					& "&amp;"
		End If

		If blnCategory Then
			pstrTemp = pstrTemp _
					 & "iLevel=1" _
					 & "&amp;txtsearchParamCat=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamCat) _
					 & "&amp;txtFromSearch=fromSearch"
		Else
			If Len(p_arySearchCriteria(enSearchCriteria_txtsearchParamCat)) > 0 Then
				pstrTemp = pstrTemp _
						 & "iLevel=" & p_arySearchCriteria(enSearchCriteria_iLevel) _
						 & "&amp;subcat=" & p_arySearchCriteria(enSearchCriteria_subcat) _
						 & "&amp;txtsearchParamCat=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamCat) _
						 & "&amp;txtCatName=" & p_arySearchCriteria(enSearchCriteria_txtCatName) _
						 & "&amp;txtFromSearch=fromSearch" & p_arySearchCriteria(enSearchCriteria_txtFromSearch)
			Else
				If pblnUseShortLink Then
					If Len(p_arySearchCriteria(enSearchCriteria_txtsearchParamMan)) > 0 And p_arySearchCriteria(enSearchCriteria_txtsearchParamMan) <> "ALL" Then pstrTemp = "txtsearchParamMan=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamMan)
					If Len(p_arySearchCriteria(enSearchCriteria_txtsearchParamVen)) > 0 And p_arySearchCriteria(enSearchCriteria_txtsearchParamVen) <> "ALL" Then pstrTemp = "txtsearchParamVen=" & p_arySearchCriteria(enSearchCriteria_txtsearchParamVen)
				End If
			End If
		End If
			
		SearchLinkParameters = pstrTemp
	
	End Function	'SearchLinkParameters

	'***********************************************************************************************

	Public Sub WriteSearchCriteria()
		Response.Write "<table border=""1"">"
		Response.Write "<tr><th colspan=2>Search Criteria</th></tr>"
		Response.Write "<tr><td>iLevel&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_iLevel) & "&nbsp;</td>"
		Response.Write "<tr><td>subcat&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_subcat) & "&nbsp;</td>"
		Response.Write "<tr><td>txtFromSearch&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtFromSearch) & "&nbsp;</td>"
		Response.Write "<tr><td>txtsearchParamTxt&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtsearchParamTxt) & "&nbsp;</td>"
		Response.Write "<tr><td>txtsearchParamType&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtsearchParamType) & "&nbsp;</td>"
		Response.Write "<tr><td>txtsearchParamCat&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtsearchParamCat) & "&nbsp;</td>"
		Response.Write "<tr><td>txtsearchParamMan&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtsearchParamMan) & "&nbsp;</td>"
		Response.Write "<tr><td>txtsearchParamVen&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtsearchParamVen) & "&nbsp;</td>"
		Response.Write "<tr><td>txtDateAddedStart&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtDateAddedStart) & "&nbsp;</td>"
		Response.Write "<tr><td>txtDateAddedEnd&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtDateAddedEnd) & "&nbsp;</td>"
		Response.Write "<tr><td>txtPriceStart&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtPriceStart) & "&nbsp;</td>"
		Response.Write "<tr><td>txtPriceEnd&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtPriceEnd) & "&nbsp;</td>"
		Response.Write "<tr><td>txtSale&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtSale) & "&nbsp;</td>"
		Response.Write "<tr><td>txtCatName&nbsp;</td><td>" & p_arySearchCriteria(enSearchCriteria_txtCatName) & "&nbsp;</td>"
		Response.Write "</table>"
	End Sub	'WriteSearchCriteria

	'***********************************************************************************************

End Class	'sfSearch

'***********************************************************************************************
'***********************************************************************************************

Sub initializeClsCategory(byRef objClsCategory, byRef blnAlreadyExists)

	If isObject(objClsCategory) Then
		blnAlreadyExists = True
		'Run another test as if objClsCategory is nothing it will pass the isObject Test
		On Error Resume Next
		If objClsCategory.CurrentCategoryName = "" Then
		End If
		
		If Err.number <> 0 Then
			If cblnSF5AE Then
				Set objClsCategory = New ssCategoryAE
			Else
				Set objClsCategory = New ssCategorySE
			End If
			Err.Clear
		End If
	Else
		blnAlreadyExists = False
		If cblnSF5AE Then
			Set objClsCategory = New ssCategoryAE
		Else
			Set objClsCategory = New ssCategorySE
		End If
	End If

End Sub	'initializeClsCategory

'***********************************************************************************************

Sub LoadSavedSearchCriteria
'this function assumes all variables are defined such as occurs on search_results.asp
'Error handling is specifically not implemented to avoid undefined variables resulting in unexpected results

Dim pclssfSearch

	Set pclssfSearch = New sfSearch
	
	With pclssfSearch
		.LoadSavedSearchCriteria("searchCriteria")
		If cblnDebugCategorySearchToolAEAddon Then
			Response.Write "<h4>Loading search criteria from cache . . .</h4>"
			.WriteSearchCriteria
		End If
		iLevel = .iLevel
		sSubcat = .subcat
		txtFromSearch = .txtFromSearch
		txtsearchParamTxt = .txtsearchParamTxt
		txtsearchParamType = .txtsearchParamType
		txtsearchParamCat = .txtsearchParamCat
		txtsearchParamMan = .txtsearchParamMan
		txtsearchParamVen = .txtsearchParamVen
		txtDateAddedStart = .txtDateAddedStart
		txtDateAddedEnd = .txtDateAddedEnd
		txtPriceStart = .txtPriceStart
		txtPriceEnd = .txtPriceEnd
		txtSale = .txtSale
	End With
	
	Set pclssfSearch = Nothing

End Sub	'LoadSavedSearchCriteria

'***********************************************************************************************

Sub releaseClsCategory(byRef objClsCategory, byRef blnAlreadyExists)
	If Not blnAlreadyExists Then
		Set objClsCategory = Nothing
		objClsCategory = Null
		Err.Clear
	End If
End Sub	'releaseClsCategory

'***********************************************************************************************

Sub SaveSearchCriteria
'this function assumes all variables are defined such as occurs on search_results.asp
'Error handling is specifically not implemented to avoid undefined variables resulting in unexpected results

Dim pclssfSearch

	Set pclssfSearch = New sfSearch
	
	With pclssfSearch
		.txtsearchParamCat = txtsearchParamCat
		.subcat = sSubcat
		.iLevel = iLevel
		.txtFromSearch = txtFromSearch
		.txtsearchParamTxt = txtsearchParamTxt
		.txtsearchParamType = txtsearchParamType
		.txtsearchParamMan = txtsearchParamMan
		.txtsearchParamVen = txtsearchParamVen
		.txtDateAddedStart = txtDateAddedStart
		.txtDateAddedEnd = txtDateAddedEnd
		.txtPriceStart = txtPriceStart
		.txtPriceEnd = txtPriceEnd
		.txtSale = txtSale
		.SaveSearchCriteria("searchCriteria")
		If cblnDebugCategorySearchToolAEAddon Then Response.Write "<h4>Saving search criteria to cache . . .</h4>"
		If cblnDebugCategorySearchToolAEAddon Then .WriteSearchCriteria
	End With
	
	Set pclssfSearch = Nothing

End Sub	'SaveSearchCriteria

'***********************************************************************************************

Sub writeCategorySelect(byVal strSubCat)

	If cblnSF5AE Then
		Call WriteSingleSelectAE(strSubCat)
	Else
		Call WriteSingleSelect()
	End If
	
End Sub	'writeCategorySelect

'***********************************************************************************************

Sub writeCategoryTrail(ByVal lngCurrentID, ByVal iLevel, ByVal sSubcat, byRef objclsCategory, byVal blnShowHome)
	If cblnSF5AE Then
		Call writeCategoryTrailAE(lngCurrentID, iLevel, sSubcat, objclsCategory, blnShowHome)
	Else
		Call writeCategoryTrailSE(txtsearchParamCat, txtsearchParamMan, txtsearchParamVen, True)
	End If
End Sub	'writeCategoryTrail

'***********************************************************************************************

Sub writeDetailCategoryTrail(byVal strProductId, byVal blnShowHome)
	If cblnSF5AE Then
		Call writeDetailCategoryTrailAE(strProductId, blnShowHome)
	Else
		Call writeCategoryTrailSE(getProductInfo(txtProdId, enProduct_CategoryID), getProductInfo(txtProdId, enProduct_MfgID), getProductInfo(txtProdId, enProduct_VendorID), blnShowHome)
	End If
End Sub	'writeDetailCategoryTrail

'***********************************************************************************************

Sub WriteSearchForm

Dim pblnAlreadyExists

	Call initializeClsCategory(mclsCategory, pblnAlreadyExists)
	If cblnSF5AE Then
		mclsCategory.WriteSearchForm
	Else
		mclsCategory.LoadCategoryScript
		mclsCategory.WriteSearchForm
	End If
	Call releaseClsCategory(mclsCategory, pblnAlreadyExists)

End Sub	'WriteSearchForm

'-----------------------------------------------------------

Sub WriteCategories(strCurrentCategory, strCurrentSubCategory, strCurrentSubSubCategory)

Dim pblnAlreadyExists

	Call initializeClsCategory(mclsCategory, pblnAlreadyExists)

	If cblnSF5AE Then
		'strCurrentCategory --> cat
		'strCurrentSubCategory --> subcat
		'strCurrentSubSubCategory --> depth
		With mclsCategory
			.DisplayCurrentLevelOnly = Not cblnExpandCategoriesByDefault
			
			.DisplayCurrentCategory = True
			.DisplayAllCategoriesAtCurrentLevel = True
			
			If Len(strCurrentSubSubCategory) > 0 And isNumeric(strCurrentSubSubCategory) Then
				If strCurrentSubSubCategory > 1 Then
					.SubcategoryFilter = strCurrentSubCategory
				Else
					.CategoryFilter = strCurrentCategory
				End If
			End If
			'.UseShortLink = True
			.WriteCategoriesAsTable
			'.WriteCategoriesAsLinks
			'.WriteCategoriesAsList
		End With
	Else
		With mclsCategory
			.CurrentCategory = strCurrentCategory
			.CurrentSubCategory = strCurrentSubCategory
			.CurrentSubSubCategory = strCurrentSubSubCategory
			.DisplayTopCategoriesOnly = Not cblnExpandCategoriesByDefault			'This is to only display top level categories due to performance issues with a large catalogs
			.DisplayCurrentCategoryOnly = False

			.WriteCategories
		End With
	End If
	Call releaseClsCategory(mclsCategory, pblnAlreadyExists)
	
End Sub	'WriteCategories

'-----------------------------------------------------------

Sub WriteProducts

Dim pblnAlreadyExists

	Call initializeClsCategory(mclsCategory, pblnAlreadyExists)

	If cblnSF5AE Then
		'mclsCategory.WriteCategoriesAsList
		'Response.Write "<hr />"
		mclsCategory.WriteProducts "", True
	Else
		mclsCategory.WriteProducts
	End If
	Call releaseClsCategory(mclsCategory, pblnAlreadyExists)
	
End Sub	'WriteProducts

'-----------------------------------------------------------

Sub WriteManufacturers(byVal lngPage, byVal lngPagesToUse)

	'Const cstrURLTemplate = "<a href=""search_results.asp?{QueryString}"" class=""clsMenuNavigationCat"" title=""{Title}"">{Text}</a>"
	Const cstrHrefTemplate = "search_results_manufacturer.asp?{QueryString}"
	Const cstrURLTemplate = "<a href=""{href}"" class=""clsMenuNavigationMfg"" title=""{Title}"">{Text}</a>"

	Dim i
	Dim pclsSFSearch
	Dim plngNumItems
	Dim pstrLineOut
	Dim pstrLink
	Dim pstrTemp
	
	Dim plngStart
	Dim plngEnd
	Dim plngPageSize
	
		'Name: 1
		'Notes: 2
		'HttpAddr: 3
		'MetaTitle: 4
		'MetaDescription: 5
		'MetaKeywords: 6
		'Description: 7
		'IsContentPage: 8

		'pstrTemp = pstrTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf
		pstrTemp = ""

		If loadManufacturerArray Then
			plngNumItems = UBound(maryManufacturers)
			If lngPagesToUse <= 1 Then
				plngStart = 0
				plngEnd = plngNumItems
			Else
				If lngPage <= 0 Then lngPage = 1
				If lngPage > lngPagesToUse Then lngPage = lngPagesToUse
				
				If ((plngNumItems + 1) Mod lngPagesToUse) = 0 Then
					plngPageSize = (plngNumItems + 1) / lngPagesToUse
				Else
					plngPageSize = Fix((plngNumItems + 1) / lngPagesToUse) + 1
				End If

				plngStart = (lngPage - 1) * plngPageSize
				plngEnd = lngPage * plngPageSize - 1
				If plngEnd > plngNumItems Then plngEnd = plngNumItems
				
			End If

			'Response.Write "UBound(maryManufacturers): " & UBound(maryManufacturers) & "<br />"
			'Response.Write "plngStart: " & plngStart & "<br />"
			'Response.Write "plngEnd: " & plngEnd & "<br />"
			'Response.Flush
			For i = plngStart To plngEnd
				pstrLineOut = ""
				'(8) isCMS - should always display
				'(9) hasCMS entry - should always display if false
				If maryManufacturers(i)(0) <> 1 And Not maryManufacturers(i)(9) Or maryManufacturers(i)(8) Then	'Exclude No Manufacturer
				'If maryManufacturers(i)(0) <> 1 Or maryManufacturers(i)(8) Then	'Exclude No Manufacturer
					'Response.Write maryManufacturers(i)(1) & "<br />"
					
					If Len(maryManufacturers(i)(3)) = 0 Then
						Set pclsSFSearch = New sfSearch
						pclsSFSearch.UseShortLink = True
						pclsSFSearch.txtsearchParamMan = maryManufacturers(i)(0)
						pstrLink = Replace(cstrHrefTemplate, "{QueryString}", pclssfSearch.SearchLinkParameters(False))
						Set pclsSFSearch = Nothing
					Else
						pstrLink = maryManufacturers(i)(3)
					End If
					pstrLink = Replace(cstrURLTemplate, "{href}", pstrLink)
					pstrLink = Replace(pstrLink, "{Title}", "")
					pstrLink = Replace(pstrLink, "{Text}", Server.HTMLEncode(maryManufacturers(i)(1)))
					'pstrLineOut = "<tr><td class=""clsMenuNavigationCat"" valign=""middle"">" & pstrLink & "</td></tr>" & vbcrlf
					pstrLineOut = pstrLink & vbcrlf
					pstrTemp = pstrTemp & pstrLineOut
				End If
			Next 'i
		End If	'loadManufacturerArray
		'pstrTemp = pstrTemp & "</table>" & vbcrlf
		
		Response.Write pstrTemp
		'Response.Flush

End Sub	'WriteManufacturers

'-----------------------------------------------------------

Sub WriteVendors

	'Const cstrURLTemplate = "<a href=""search_results.asp?{QueryString}"" class=""clsMenuNavigationCat"" title=""{Title}"">{Text}</a>"
	Const cstrURLTemplate = "<a href=""search_results_manufacturer.asp?{QueryString}"" class=""clsMenuNavigationCat"" title=""{Title}"">{Text}</a>"

	Dim i
	Dim pclsSFSearch
	Dim plngNumItems
	Dim pstrLineOut
	Dim pstrLink
	Dim pstrTemp
	
		pstrTemp = pstrTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf

		If loadVendorArray Then
			plngNumItems = UBound(maryVendors)
			For i = 0 To plngNumItems
				pstrLineOut = ""
				If maryVendors(i)(0) <> 1 Then	'Exclude No Vendor
					'Response.Write maryVendors(i)(1) & "<br />"
					
					Set pclsSFSearch = New sfSearch
					pclsSFSearch.UseShortLink = True
					pclsSFSearch.txtsearchParamVen = maryVendors(i)(0)
					pstrLink = Replace(cstrURLTemplate, "{QueryString}", pclssfSearch.SearchLinkParameters(False))
					pstrLink = Replace(pstrLink, "{Title}", "")
					pstrLink = Replace(pstrLink, "{Text}", Server.HTMLEncode(maryVendors(i)(1)))
					Set pclsSFSearch = Nothing
					pstrLineOut = "<tr><td class=""clsMenuNavigationCat"" valign=""middle"">" & pstrLink & "</td></tr>" & vbcrlf
					pstrTemp = pstrTemp & pstrLineOut
				End If
			Next 'i
		End If	'loadVendorArray
		pstrTemp = pstrTemp & "</table>" & vbcrlf
		
		Response.Write pstrTemp
		'Response.Flush

End Sub	'WriteVendors

'-----------------------------------------------------------

Sub WriteSingleSelect

Dim pblnAlreadyExists

	Call initializeClsCategory(mclsCategory, pblnAlreadyExists)
	If cblnSF5AE Then
		Response.Write mclsCategory.SelectCategoryAE("", 1)
		Response.Write  WriteHiddenField("txtsearchParamCat","ALL")
		Response.Write  WriteHiddenField("subcat","")
		Response.Write  WriteHiddenField("txtCatName","")
	Else
		mclsCategory.WriteSingleSelect
	End If
	
End Sub	'WriteSingleSelect

%>