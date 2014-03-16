<%
'********************************************************************************
'*   Search Grid for StoreFront 5.0 					                        *
'*   Release Version:	1.2.1                                                   *
'*   Release Date:      February 10, 2002		                                *
'*   Revision Date:		May 2, 2003												*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   1.2.1 - added support for a spacer row between rows						*
'*                                                                              *
'*   1.2 - added support for ordering from search grid							*
'*         added option to disable ordering for items with attributes           *
'*                                                                              *
'*   1.1 - added check for null value in sale price/sale is active              *
'*         SF doesn't always set a default value as it should                   *
'*         Moved cell alignment settings to user configuration section          *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Dim mbytCurrentColumn
Dim mstrSpacerRow

mbytCurrentColumn = 0

	'////////////////////////////////////////////////////////////////////////////////
	'//
	'//		USER CONFIGURATION

		Const cbytNumColumns = 2

		'Display Contents
		'if you want to use a "Image Not Available" image just set the path below
		'leave it blank to display Image Not Available
		'Const cstrImageNotAvailablePath = "images/NoImage.gif"
		Const cstrImageNotAvailablePath = ""
		
		Const cblnDisplayInventoryStatus = False
		Const cblnDisplayMTPStatus = False
		
		Const cblnAllowOrdering = True
		Const cblnShowIndividualAddToCartButtons = True
		Const cblnDisableOrderingIfAttributes = True
		Const cstrQty_AddToCartButtonSeparator = "<br />"	'html to appear between qty box and add to cart button; usually will be "<br />" or "&nbsp;"
		Const cstrShowMoreDetails = "More Details"	'Leave blank to not show
		'Const cstrShowMoreDetails = ""	'Leave blank to not show
		
		mstrSpacerRow = "<tr><td colspan=" & cbytNumColumns & "><br /><hr /><br /></td></tr>" & vbcrlf

	'//
	'//
	'////////////////////////////////////////////////////////////////////////////////

Sub WritessSearchGrid

Dim i
Dim pstrHREF
Dim pstrImage
Dim pstrProductName
Dim pstrProductDescription
Dim pstrProductPrice
Dim pstrAttributes

'Array decoder
' 0- sfProducts.ProdID
' 1- sfProducts.prodName
' 2- sfProducts.prodImageSmallPath
' 3- sfProducts.prodLink
' 4- sfProducts.prodPrice
' 5- sfProducts.prodSaleIsActive
' 6- sfProducts.prodSalePrice
' 7- sfProducts.prodDescription
' 8- sfProducts.prodAttrNum
' 9- sfProducts.prodCategoryId
' 10- sfProducts.prodShortDescription
' 11- sfProducts.prodPLPrice
' 12- sfProducts.prodPLSalePrice
' 13- sfVendors.vendName
' 14- sfManufacturers.mfgName
 
    If iRec=0 Then Response.Write "<tr><td align=""center""><table class=""searchGridTable"">" & vbcrlf

	If mbytCurrentColumn = 0 Then Response.Write "<tr>" & vbcrlf
	mbytCurrentColumn = mbytCurrentColumn + 1
	
	Response.Write "<td class=""searchGridCell"" width=""" & FormatPercent(1/cbytNumColumns,0) & """>" & vbcrlf

	'Create the link - modified to work with SEOptimizer
	If (Len(Trim(arrProduct(3, iRec))) > 0) Then
		pstrHREF = "<a href=""" & arrProduct(3, iRec) & """>"
	Else
		pstrHREF = "<a href=""" & "detail.asp?PRODUCT_ID=" & arrProduct(0, iRec) & """>"
	End If
		
	'Write the Image link
	If (Len(txtImagePath) > 0) Then
		pstrImage = "  " & pstrHREF & "<img class=""inputImage"" src=""" & txtImagePath & """ alt=""" & Server.HTMLEncode(StripHTML(arrProduct(1, iRec))) & """></a>" & vbcrlf
	Else
		If Len(cstrImageNotAvailablePath) > 0 Then
			pstrImage = "  " & pstrHREF & "<img class=""inputImage"" src=""" & cstrImageNotAvailablePath & """ alt=""" & "No Image Available" & """></a>" & vbcrlf
		Else
			pstrImage = "  " & pstrHREF & "<i>No Image Available</i></a>" & vbcrlf
		End If
	End If

	'Write the text link
	pstrProductName = pstrHREF & "<span class=""productName"">" & arrProduct(1, iRec) & "</span></a>" & vbcrlf
	pstrProductDescription = arrProduct(10, iRec) & vbcrlf

	'Fix if null in on sale
	If isNull(arrProduct(5, iRec)) Then arrProduct(5, iRec) = 0

	'Create the price
	If cBool(arrProduct(5, iRec)) Then 'is product on sale
		If IsNull(FormatCurrency(arrProduct(4, iRec))) Then
			pstrProductPrice = "<strike>Please contact customer service</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		Else
			pstrProductPrice = "<strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		End If
	Else
		pstrProductPrice = FormatCurrency(arrProduct(4, iRec)) & vbcrlf
	End If 
	
	If cblnAllowOrdering Then
		If irsSearchAttRecordCount <> "" Then
			For iAttCounter = 0 to irsSearchAttRecordCount
				If arrProduct(0, iRec) = arrAtt(2, iAttCounter) Then
					pstrAttributes = pstrAttributes & "<br /><FONT face='" & C_FONTFACE4 & "' color='" &  C_FONTCOLOR4 & "' SIZE='" &  C_FONTSIZE4 & "'>" & arrAtt(1, iAttCounter) & "</FONT>" & vbcrlf
					pstrAttributes = pstrAttributes & "  <SELECT size='1' name='attr" & icounter & "." & arrProduct(0, iRec) & "' style='" & C_FORMDESIGN & "'>" & vbcrlf
					For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
						If isArray(arrAttDetail) Then
						If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
							sAmount = ""
							Select Case arrAttDetail(4, iAttDetailCounter)
								Case 1 
									sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
								Case 2 
									sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
							End Select
							pstrAttributes = pstrAttributes &  "    <option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>" & vbcrlf
						End If
						End If
					Next
					pstrAttributes = pstrAttributes & "</SELECT>" & vbcrlf
				End If
				icounter = icounter + 1
			Next
		End If 
	End If	'cblnAllowOrdering
	
	'Now Write the contents of the cell
	'Disable ordering if attributes are present
	Dim pblnDisableOrdering:	pblnDisableOrdering = False

	If cblnDisableOrderingIfAttributes Then	pblnDisableOrdering = CBool(Len(pstrAttributes) > 0)

	If Not cblnAllowOrdering Then pblnDisableOrdering = True

	Response.Write pstrImage
	Response.Write "<br />" & pstrProductName
	Response.Write "<div class=""searchGridProductID"">" & arrProduct(0, iRec) & "</div>"
	If Len(pstrProductDescription) > 0 Then Response.Write "<div class=""searchGridProductDescription"">" & pstrProductDescription & "</div>"
	If Not pblnDisableOrdering Then Response.Write pstrAttributes & vbcrlf
	Response.Write "<div class=""sellPrice"">" & pstrProductPrice & "</div>" & vbcrlf
	
	If cblnSF5AE And cblnDisplayInventoryStatus Then SearchResults_GetProductInventory arrProduct(0, iRec)
	If cblnSF5AE And cblnDisplayMTPStatus Then SearchResults_ShowMTPricesLink arrProduct(0, iRec)

	If Not pblnDisableOrdering Then Response.Write "Quantity:&nbsp;<input style='" & C_FORMDESIGN & "'  type='text' name='QUANTITY." & arrProduct(0, iRec) & "' title='Quantity' size='3' value='' onblur='return isInteger(this, true, " & Chr(34) & "Please enter an integer greater than one for the quantity" & Chr(34) & ")'>"
	If Not pblnDisableOrdering And cblnShowIndividualAddToCartButtons Then Response.Write cstrQty_AddToCartButtonSeparator & "<input type='image' name='AddProduct' class=""inputImage"" src='" & C_BTN03 & "' alt='Add To Cart'><br />"
	If Len(cstrShowMoreDetails) > 0 Then Response.Write "<div class=""searchGridMoreDetails"">" & pstrHREF & cstrShowMoreDetails & "</a></div>"
	
	Response.Write "</td>" & vbcrlf
	
	If mbytCurrentColumn = cbytNumColumns Then
		Response.Write "</tr>" & vbcrlf
		If Len(mstrSpacerRow) > 0 Then Response.Write mstrSpacerRow
		mbytCurrentColumn = 0
	End If
	
	If (iRec=iVarPageSize-1) Then
		If mbytCurrentColumn > 0 Then
			For i=1 to (cbytNumColumns - mbytCurrentColumn)
				Response.Write "<td>&nbsp;</td>" & vbcrlf
			Next
		End If
		Response.Write "</table></td></tr>" & vbcrlf
	End If
	
End Sub	'WritessSearchGrid

'***************************************************************************************************************************************************************************

Sub WritessTabularSearchResults

Dim plngPos
Dim pstrHREF
Dim pstrImage
Dim pstrProductName
Dim pstrProductPrice
Dim pstrAttributes
Dim plngCellCount

    If iRec=0 Then
		Response.Write "<center>" & vbcrlf
		Response.Write "<table cellpadding=""6"" cellspacing=""0"" border=""1"" bordercolor=""black"" style=""border-collapse:collapse;"">" & vbcrlf
		'Response.Write "<table class=""searchGridTable"">" & vbcrlf
		Response.Write "  <colgroup>" & vbcrlf
		Response.Write "    <col align='left' valign='top' />" & vbcrlf
		Response.Write "    <col align='left' valign='top' />" & vbcrlf
		Response.Write "    <col align='left' valign='top' />" & vbcrlf
		Response.Write "    <col align='left' valign='top' />" & vbcrlf
		Response.Write "  </colgroup>" & vbcrlf

		'Write header row
		Response.Write "  <tr>" & vbcrlf
		Response.Write "    <th class='searchGridHeader'>&nbsp;Product Name</th>" & vbcrlf
		Response.Write "    <th class='searchGridHeader'>&nbsp;Overview</th>" & vbcrlf
		Response.Write "    <th class='searchGridHeader'>&nbsp;Our Price</th>" & vbcrlf
		Response.Write "    <th class='searchGridHeader'>&nbsp;Summary</th>" & vbcrlf
		Response.Write "  </tr>" & vbcrlf

	End If

	'Create the link
	If (Len(Trim(arrProduct(3, iRec))) > 0) Then
		pstrHREF = arrProduct(3, iRec)
	Else
		pstrHREF = "detail.asp?PRODUCT_ID=" & arrProduct(0, iRec)
	End If
	
	'Write the Image link
	If (Len(txtImagePath) > 0) Then
		pstrImage = "<img border=""0"" src=""" & txtImagePath & """ alt=""" & Server.HTMLEncode(StripHTML(arrProduct(1, iRec))) & """>"
	Else
		If Len(cstrImageNotAvailablePath) > 0 Then
			pstrImage = "<img border=""0"" src=""" & cstrImageNotAvailablePath & """ alt=""" & "No Image Available" & """>"
		Else
			pstrImage = "<i>No Image Available</i>"
		End If
	End If

	'Fix if null in on sale
	If isNull(arrProduct(5, iRec)) Then arrProduct(5, iRec) = 0

	'Create the price
	If cBool(arrProduct(5, iRec)) Then 'is product on sale
		If IsNull(FormatCurrency(arrProduct(4, iRec))) Then
			pstrProductPrice = "<strike>Please contact customer service</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		Else
			pstrProductPrice = "<strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		End If
	Else
		pstrProductPrice = FormatCurrency(arrProduct(4, iRec)) & vbcrlf
	End If 
	
	icounter = 0
	If True Then
		If irsSearchAttRecordCount <> "" Then
			For iAttCounter = 0 to irsSearchAttRecordCount
				If arrProduct(0, iRec) = arrAtt(2, iAttCounter) Then
					icounter = icounter + 1
					pstrAttributes = pstrAttributes & "<br /><FONT face='" & C_FONTFACE4 & "' color='" &  C_FONTCOLOR4 & "' SIZE='" &  C_FONTSIZE4 & "'>" & arrAtt(1, iAttCounter) & "</FONT>" & vbcrlf
					'pstrAttributes = pstrAttributes & "  <SELECT size='1' name='attr" & icounter & cstrSSMPOAttributeDelimiter & arrProduct(0, iRec) & cstrSSTextBasedAttributeHTMLDelimiter & "' style='" & C_FORMDESIGN & "'>" & vbcrlf
					pstrAttributes = pstrAttributes & "  " & DisplayAttributesSearchResults("frmSearchResults")
					'For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					'	If isArray(arrAttDetail) Then
					'	If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
					'		sAmount = ""
					'		Select Case arrAttDetail(4, iAttDetailCounter)
					'			Case 1 
					'				sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
					'			Case 2 
					'				sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
					'		End Select
					'		pstrAttributes = pstrAttributes &  "    <option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>" & vbcrlf
					'	End If
					'	End If
					'Next
					pstrAttributes = pstrAttributes & "</SELECT>" & vbcrlf
				End If
			Next
		End If 
	End If	'cblnAllowOrdering
	
	'Now Write the contents of the row
	Response.Write "<tr>"
	
	Response.Write "<td>"
	'Response.Write "<div class=""searchGridProductID"">" & arrProduct(0, iRec) & "</div>"
	Response.Write "<div align=""center""><a href=""" & pstrHREF & """ title=""" & Server.HTMLEncode(StripHTML(arrProduct(10, iRec))) & """>" & pstrImage & "</a></div>"
	Response.Write "</td>"
	
	Response.Write "<td>"
	Response.Write "<a href=""" & pstrHREF & """ title=""" & Server.HTMLEncode(StripHTML(arrProduct(10, iRec))) & """>" & arrProduct(1, iRec) & "</a>"
	Response.Write "<p><div class=""searchGridProductDescription"">" & arrProduct(10, iRec) & "</div></p>" 
	'Response.Write "<p class=""searchGridProductDescription"">Click the image at left for a more detailed description</p>" 
	Response.Write "</td>"
	
	Response.Write "<td>"
	Response.Write "<div class=""sellPrice"">" & pstrProductPrice & "</div>" 
	Response.Write pstrAttributes & vbcrlf
	Response.Write "Quantity:&nbsp;<input type=""text"" class=""formDesign"" name=""QUANTITY." & arrProduct(0, iRec) & """ id=""QUANTITY." & arrProduct(0, iRec) & """ title=""Quantity"" size=""3"" value="""" onblur='return isInteger(this, true, " & Chr(34) & "Please enter an integer greater than one for the quantity" & Chr(34) & ")'><br />"
	Response.Write "<input type='image' name='AddProduct' border='0' src='" & C_BTN03 & "' alt='Add To Cart'>"
	Response.Write "</td>"
	
	'Set maximum length to display for description
	If clngMaxLengthDescription = 0 Then
	ElseIf clngMaxLengthDescription <> -1 And Len(arrProduct(7, iRec)) > clngMaxLengthDescription Then
		plngPos = InStrRev(arrProduct(7, iRec), " ", clngMaxLengthDescription)

		'account for long initial sentences
		If plngPos > 0 Then
			arrProduct(7, iRec) = Left(arrProduct(7, iRec), plngPos)
		Else
			plngPos = InStr(clngMaxLengthDescription, arrProduct(7, iRec), " ")
			'plngPos = InStrRev(arrProduct(7, iRec), "<br", clngMaxLengthDescription)
			arrProduct(7, iRec) = Left(arrProduct(7, iRec), plngPos)
		End If

	End If
	Response.Write "<td>"
	If clngMaxLengthDescription > -1 Then Response.Write "<div class=""searchGridProductDescription"">" & arrProduct(7, iRec) & "</div><br />" 
	Response.Write "<div class=""searchGridProductDescription"" align=""center"">(<a href=""" & pstrHREF & """ title=""" & Server.HTMLEncode(StripHTML(arrProduct(1, iRec))) & """>Read full description</a>)</div>" 
	Response.Write "</td>"
	
	Response.Write "</tr>"

	If (iRec=iVarPageSize-1) Then Response.Write "</table></center>" & vbcrlf
	
End Sub	'WritessTabularSearchResults

'***************************************************************************************************************************************************************************

Sub WritessTabularSearchResults_v0

Dim i
Dim pstrHREF
Dim pstrImage
Dim pstrProductName
Dim pstrProductPrice
Dim pstrAttributes
Dim paryShowColumn(6)
Dim plngCellCount
Const cblnCombineNameDescriptionInSameCell = True
Const cblnShowSpacerHR = True

	'Format for the table
	paryShowColumn(0) = False	'show small image
	paryShowColumn(1) = False	'show product id
	paryShowColumn(2) = True	'show product name and short description
	paryShowColumn(3) = True	'show product price
	paryShowColumn(4) = False	'show attributes
	paryShowColumn(5) = False	'show qty
	paryShowColumn(6) = False	'show add to cart/saved cart/email friend buttons
	
	If cblnShowSpacerHR Then
		plngCellCount = 0
		For i = 0 To UBound(paryShowColumn)
			If paryShowColumn(i) Then plngCellCount = plngCellCount + 1
			If Not cblnCombineNameDescriptionInSameCell Then plngCellCount = plngCellCount + 1
		Next 'i
	End If

    If iRec=0 Then
		Response.Write "<center>" & vbcrlf
		Response.Write "<table border='0' width='95%' background='' cellpadding='3' cellspacing='0'>" & vbcrlf
		If paryShowColumn(0) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		If paryShowColumn(1) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		If cblnCombineNameDescriptionInSameCell Then
			If paryShowColumn(2) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		Else
			If paryShowColumn(2) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
			If paryShowColumn(2) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		End If
		If paryShowColumn(3) Then Response.Write "  <colgroup align='right' valign='top'>" & vbcrlf
		If paryShowColumn(4) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		If paryShowColumn(5) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf
		If paryShowColumn(6) Then Response.Write "  <colgroup align='left' valign='top'>" & vbcrlf

		'Write header row
		Response.Write "  <tr>" & vbcrlf
		If paryShowColumn(0) Then Response.Write "    <th class='clsMenuHeader'>&nbsp;</th>" & vbcrlf
		If paryShowColumn(1) Then Response.Write "    <th class='clsMenuHeader' align='left'>Product ID</th>" & vbcrlf
		If cblnCombineNameDescriptionInSameCell Then
			If paryShowColumn(2) Then Response.Write "    <th class='clsMenuHeader' align='left'>Product</th>" & vbcrlf
		Else
			If paryShowColumn(2) Then Response.Write "    <th class='clsMenuHeader' colspan='2' align='left'>Product</th>" & vbcrlf
		End If
		If paryShowColumn(3) Then Response.Write "    <th class='clsMenuHeader'>Price</th>" & vbcrlf
		If paryShowColumn(4) Then Response.Write "    <th class='clsMenuHeader'>&nbsp;</th>" & vbcrlf
		If paryShowColumn(5) Then Response.Write "    <th class='clsMenuHeader'>Qty&nbsp;</th>" & vbcrlf
		If paryShowColumn(6) Then Response.Write "    <th class='clsMenuHeader'>&nbsp;</th>" & vbcrlf
		Response.Write "  </tr>" & vbcrlf

		'Write an empty row
		Response.Write "    <tr><td align='center' colspan='" & plngCellCount & "'>&nbsp;</td></tr>" & vbcrlf
	End If

	'Create the link
	If (Len(Trim(arrProduct(3, iRec))) > 0) Then
		pstrHREF = "<a href=""" & arrProduct(3, iRec) & """>"
	Else
		pstrHREF = "<a href=""" & "detail.asp?PRODUCT_ID=" & arrProduct(0, iRec) & """>"
	End If
	
	'Write the Image link
	If (Len(txtImagePath) > 0) Then
		pstrImage = "  " & pstrHREF & "<img border=""0"" src=""" & txtImagePath & """ alt=""" & Server.HTMLEncode(StripHTML(arrProduct(0, iRec))) & """></a>" & vbcrlf
	Else
		If Len(cstrImageNotAvailablePath) > 0 Then
			pstrImage = "  " & pstrHREF & "<img border=""0"" src=""" & cstrImageNotAvailablePath & """ alt=""" & "No Image Available" & """></a>" & vbcrlf
		Else
			pstrImage = "  " & pstrHREF & "<i>No Image Available</i></a>" & vbcrlf
		End If
	End If

	'Write the text link
	pstrProductName = pstrHREF & arrProduct(1, iRec) & "</a>" & vbcrlf

	'Fix if null in on sale
	If isNull(arrProduct(5, iRec)) Then arrProduct(5, iRec) = 0

	'Create the price
	If cBool(arrProduct(5, iRec)) Then 'is product on sale
		If IsNull(FormatCurrency(arrProduct(4, iRec))) Then
			pstrProductPrice = "  <br /><strike>Please contact customer service</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		Else
			pstrProductPrice = "<strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike><br /><font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		End If
	Else
		pstrProductPrice = FormatCurrency(arrProduct(4, iRec)) & vbcrlf
	End If 
	
'Now Write the contents of the cell

	Response.Write "  <tr>" & vbcrlf
	If paryShowColumn(0) Then Response.Write "    <td>" & pstrImage & "</td>" & vbcrlf
	If paryShowColumn(1) Then Response.Write "    <td><a href=""" & arrProduct(3, iRec) & """>" & arrProduct(0, iRec) & "</a></td>" & vbcrlf
	If cblnCombineNameDescriptionInSameCell Then
		If paryShowColumn(2) Then Response.Write "    <td><b>" & pstrProductName & "</b>&nbsp;-&nbsp;" & arrProduct(11, iRec) & "</td>" & vbcrlf
	Else
		If paryShowColumn(2) Then Response.Write "    <td><b>" & pstrProductName & "</b></td>" & vbcrlf
		If paryShowColumn(2) Then Response.Write "    <td>" & arrProduct(11, iRec) & "</td>" & vbcrlf
	End If
	If paryShowColumn(3) Then Response.Write "    <td>" & pstrProductPrice & "</td>" & vbcrlf
	If paryShowColumn(4) Then
		Response.Write "    <td>"
		Call DisplayAttributes_TabularSearchResults
		Response.Write "</td>" & vbcrlf
	End If
	If paryShowColumn(5) Then Response.Write "    <td>" & "<INPUT style=" & Chr(34) & C_FORMDESIGN & Chr(34) & "  type=""text"" name=""QUANTITY." & arrProduct(0, iRec) & Chr(34) & " id=""QUANTITY." & arrProduct(0, iRec) & Chr(34) & " title=""Quantity"" size=""3"" value="""" onblur=""return isInteger(this, true, 'Please enter an integer greater than one for the quantity')""></td>" & vbcrlf
	If paryShowColumn(6) Then Response.Write "    <td>" & "<input type=""image"" name=""AddProduct"" border=""0"" src=""" & C_BTN03 & """ alt=""Add To Cart"">" & "</td>" & vbcrlf
	Response.Write "  </tr>" & vbcrlf

	If (iRec=iVarPageSize-1) Then 
		Response.Write "</table></center>" & vbcrlf
	Else
		If cblnShowSpacerHR Then
			Response.Write "    <tr><td align='center' colspan='" & plngCellCount & "'><hr color='#6699CC' width='100%'></td></tr>" & vbcrlf
		End If
	End If
	
End Sub	'WritessTabularSearchResults_v0

	'**********************************************************************************************************

	Sub DisplayAttributes_TabularSearchResults()
%>
                      <TABLE border="0" align="center">
                        <%
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: BEGIN --
                            ' -------------------------------------------
                            If irsSearchAttRecordCount <> "" Then
								For iAttCounter = 0 to irsSearchAttRecordCount
									If arrProduct(0, iRec) = arrAtt(2, iAttCounter) Then
%>                    
                        <TR>                
                          <TD align="right"><%= arrAtt(1, iAttCounter) %></TD>
                          <TD><SELECT size="1" name="attr<%= icounter & "." & arrProduct(0, iRec) %>" ID="attr<%= icounter & "." & arrProduct(0, iRec) %>" class="formDesign">
                              <%
										For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
											If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
														sAmount = ""
												Select Case arrAttDetail(4, iAttDetailCounter)
													Case 1 
														sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
													Case 2 
														sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
												End Select
												Response.Write "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>"
											End If
										Next
%>
                            </SELECT></TD>
                        </TR>
                        <%  
									icounter = icounter + 1
									End If 
								Next
							End If 
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: END
                            ' -------------------------------------------
                     
%>
                      </TABLE>
<%		
	End Sub	'DisplayAttributes_TabularSearchResults

%>