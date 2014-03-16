<%

'***********************************************************************************************
'***********************************************************************************************

Class ssCategoryFragment

Private pobjRSCategories

Private plngCategoryID
Private pstrCategoryName
Private pstrCategoryImage
Private pstrCategoryDescription
Private paryAlphabeticalSubCategories
Private paryRandomSubCategories
Private plngNumSubCategories

Private cstrCategoryLink
Private cstrSubCategoryLink
Private cstrSubCategoryLink_Alt

Private plngCatID
Private plngSubCatID
Private plngDepth

	'***********************************************************************************************

	Private Sub class_Terminate()
		On Error Resume Next
		pobjRSCategories.Close
		Set pobjRSCategories = nothing
	End Sub
	Private Sub class_Initialize()
		cstrCategoryLink =			"search_results.asp?txtsearchParamType=ALL&txtsearchParamMan=ALL&txtsearchParamVen=ALL&iLevel=1&txtsearchParamCat={CatID}"
		cstrSubCategoryLink =		"search_results.asp?txtsearchParamType=ALL&txtsearchParamMan=ALL&txtsearchParamVen=ALL&txtsearchParamTxt=*&txtFromSearch=fromSearch&txtsearchParamCat={CatID}&subcat={SubCatID}&iLevel={Depth}&txtCatName={Depth}"
		cstrSubCategoryLink_Alt =	"category_search_results.asp?Category={CatID}&txtsearchParamType=ALL&txtsearchParamMan=ALL&txtsearchParamVen=ALL&txtsearchParamTxt=*&txtFromSearch=fromSearch&txtsearchParamCat={CatID}&subcat={SubCatID}&iLevel={Depth}&txtCatName={Depth}"
	End Sub
	
	'***********************************************************************************************

	Public Property Get rsCategories
		If isObject(pobjRSCategories) Then Set rsCategories = pobjRSCategories
	End Property
	
	Public Property Get AlphabeticalSubCategories
		AlphabeticalSubCategories = paryAlphabeticalSubCategories
	End Property
	
	Public Property Get RandomSubCategories
		RandomSubCategories = paryRandomSubCategories
	End Property
	
	Public Property Get NumSubCategories
		NumSubCategories = plngNumSubCategories
	End Property
	
	Public Property Get CategoryID
		CategoryID = plngCategoryID
	End Property
	
	Public Property Get CategoryName
		CategoryName = pstrCategoryName
	End Property
	
	Public Property Get CategoryDescription
		CategoryDescription = pstrCategoryDescription
	End Property
	
	Public Property Get CategoryImage
		CategoryImage = pstrCategoryImage
	End Property

	'***********************************************************************************************

	Public Sub SetCurrentCategory(byVal lngCatID, byVal lngSubCatID, byVal lngDepth)
		plngCatID = plngCatID
		plngSubCatID = lngSubCatID
		plngDepth = lngDepth
	End Sub

	'***********************************************************************************************

	Public Function CategoryLink(byVal aryItem)
		CategoryLink = doReplacements(cstrCategoryLink, aryItem)
	End Function	'CategoryLink
	
	Public Function SubCategoryLink(byVal aryItem)
		SubCategoryLink = doReplacements(cstrSubCategoryLink, aryItem)
	End Function	'SubCategoryLink
	
	Public Function SubCategoryLink_Alt(byVal aryItem)
		SubCategoryLink_Alt = doReplacements(cstrSubCategoryLink_Alt, aryItem)
	End Function	'SubCategoryLink
	
	Private Function doReplacements(byVal strIn, byVal aryItem)
	
	Dim pstrOut
	
		pstrOut = Replace(strIn, "{CatID}", plngCategoryID)
		If isArray(aryItem) Then
			pstrOut = Replace(pstrOut, "{SubCatID}", aryItem(0))
			pstrOut = Replace(pstrOut, "{Depth}", aryItem(2)+1)
		End If
		
		doReplacements = pstrOut
	
	End Function	'doReplacements

	'***********************************************************************************************

	Private Sub setRandomSubCategories
	
	Dim i
	Dim pblnInitialized
	Dim pdicUnassignedProducts
	Dim plngUnassignedProducts
	Dim plngRandomIndex
	Dim paryKeys
	
		pblnInitialized = isArray(paryRandomSubCategories)
		If pblnInitialized Then pblnInitialized = CBool(UBound(paryAlphabeticalSubCategories) = UBound(paryRandomSubCategories))

		If Not pblnInitialized Then
			ReDim paryRandomSubCategories(plngNumSubCategories - 1)
			Set pdicUnassignedProducts = CreateObject("SCRIPTING.DICTIONARY")
			For i = 0 To plngNumSubCategories - 1
				pdicUnassignedProducts.Add i, CStr(i)
			Next 'i

			Randomize
			For i = 0 To plngNumSubCategories - 1
				plngUnassignedProducts = pdicUnassignedProducts.Count
					
				paryKeys = pdicUnassignedProducts.Keys()
				plngRandomIndex = Int(((plngUnassignedProducts) * Rnd))

				paryRandomSubCategories(i) = paryAlphabeticalSubCategories(CLng(paryKeys(plngRandomIndex)))
				pdicUnassignedProducts.Remove(paryKeys(plngRandomIndex))
			Next 'i
		End If

	End Sub	'setRandomSubCategories
	
	'***********************************************************************************************

	Private Sub LoadValues
	
	Dim i

		With pobjRSCategories
			If Not .EOF Then
				plngCategoryID = .Fields("catID").Value
				pstrCategoryName = .Fields("catName").Value
				pstrCategoryDescription = .Fields("catDescription").Value
				pstrCategoryImage = .Fields("catImage").Value
				
				plngNumSubCategories = .RecordCount
				ReDim paryAlphabeticalSubCategories(plngNumSubCategories - 1)
				For i = 0 To plngNumSubCategories - 1
					paryAlphabeticalSubCategories(i) = Array(.Fields("subcatID").Value, .Fields("subcatName").Value, .Fields("Depth").Value)
					.MoveNext
				Next 'i
				Call setRandomSubCategories
			End If
		End With

	End Sub	'LoadValues
	
	'***********************************************************************************************

	Public Function LoadCategoryFragmentByCategoryID(byVal lngCatID)

	Dim pstrSQL
	Dim pblnResult

		pblnResult = False
		
		If Len(lngCatID) = 0 OR Not isNumeric(lngCatID) Then
			LoadCategoryFragmentByCategoryID = pblnResult 
			Exit Function
		End If
		
		pstrSQL = "SELECT sfCategories.catID, sfCategories.catDescription, sfCategories.catImage, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.bottom" _
				& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
				& " Where catIsActive=1 AND (sfSub_Categories.CatHierarchy Not Like 'none-%') AND sfCategories.catID=" & lngCatID _
				& " ORDER BY sfCategories.catName, sfSub_Categories.Depth, sfSub_Categories.subcatName"

		Set pobjRSCategories = CreateObject("ADODB.RECORDSET")
		With pobjRSCategories
			.CursorLocation = 2 'adUseClient

			'On Error Resume Next
			'Response.Write "sql: " & pstrSQL & "<br />" & vbcrlf
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number <> 0 Then
				Response.Write "<font color=red>Error in LoadCategoryFragmentByCategoryID: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				Response.Write "<font color=red>Error in LoadCategoryFragmentByCategoryID: sql = " & pstrSQL & "</font><br />" & vbcrlf
				Response.Flush
				Err.Clear
			ElseIf Not .EOF Then
				Call LoadValues			
				pblnResult = True
			End If
			
		End With
		
		LoadCategoryFragmentByCategoryID = pblnResult 
			
	End Function	'LoadCategoryFragmentByCategoryID

	'***********************************************************************************************

End Class	'ssCategoryFragment

'***********************************************************************************************
'***********************************************************************************************
%>