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

		'Display Style
		Const cstrTableWidth = "100%"
		Const cstrTableBorder = "0"
		Const cstrTableCellPadding = "2"
		Const cstrTableCellSpacing = "0"
		Const cstrTableBackground = ""

		Const cstrCellAlign = "center"	'options (left, center, right)
		Const cstrCellVAlign = "top"	'options (top, middle, bottom)
		
		Const cbytNumColumns = 3

		'Display Contents
		'if you want to use a "Image Not Available" image just set the path below
		'leave it blank to display Image Not Available
		'Const cstrImageNotAvailablePath = "images/NoImage.gif"
		Const cstrImageNotAvailablePath = ""
		
		Const cblnDisplaySmallImage = True
		Const cblnDisplayProductID = True
		Const cblnDisplayShortDescription = True
		
		Const cblnAllowOrdering = True
		Const cblnShowIndividualAddToCartButtons = True
		Const cblnDisableOrderingIfAttributes = False
		Const cstrQty_AddToCartButtonSeparator = "<br />"	'html to appear between qty box and add to cart button; usually will be "<br />" or "&nbsp;"
		Const cstrShowMoreDetails = "More Details"	'Leave blank to not show
		'Const cstrShowMoreDetails = ""	'Leave blank to not show
		
		mstrSpacerRow = "<tr><td colspan=" & cbytNumColumns & "><p><hr></p></td></tr>" & vbcrlf

	'//
	'//
	'////////////////////////////////////////////////////////////////////////////////

		mstrSpacerRow = "<tr><td colspan=" & cbytNumColumns & ">" & mstrSpacerRow & "</td></tr>" & vbcrlf

Sub WritessSearchGrid

Dim i
Dim pstrHREF
Dim pstrImage
Dim pstrProductName
Dim pstrProductPrice
Dim pstrAttributes

Dim pstrProductTemplate
Dim pstrProductID
Dim pstrProductDesc
Dim pstrProductSalePrice
Dim pstrProductURL
Dim pstrProductImage

    If iRec=0 Then Response.Write "<center><table border='" & cstrTableBorder & "' width='" & cstrTableWidth & "' background='" & cstrTableBackground & "' cellpadding='" & cstrTableCellPadding & "' cellspacing='" & cstrTableCellSpacing & "'>" & vbcrlf

	If mbytCurrentColumn = 0 Then Response.Write "<tr>" & vbcrlf
	mbytCurrentColumn = mbytCurrentColumn + 1
	
	Response.Write "<td valign='" & cstrCellVAlign & "' align='" & cstrCellAlign & "' width='" & FormatPercent(1/cbytNumColumns,0) & "'>" & vbcrlf

	'Create the link - modified to work with SEOptimizer
	If (Len(Trim(arrProduct(3, iRec))) > 0) Then
		If cblnUseSEOptimizer Then
			'pstrHREF = "<a href=""" & getPageLink(arrProduct(3, iRec), arrProduct(1, iRec), arrProduct(13, iRec)) & """>"
			pstrHREF = "<a href=""" & getPageLink(arrProduct(0, iRec), arrProduct(3, iRec), arrProduct(1, iRec), arrProduct(12, iRec)) & """>"
		Else
			pstrHREF = "<a href=""" & arrProduct(3, iRec) & """>"
		End If
	Else
		If cblnUseSEOptimizer Then
			'pstrHREF = "<a href=""" & getPageLink(arrProduct(3, iRec),arrProduct(1, iRec),arrProduct(13, iRec)) & """>"
			pstrHREF = "<a href=""" & getPageLink(arrProduct(0, iRec), arrProduct(3, iRec), arrProduct(1, iRec), arrProduct(12, iRec)) & """>"
		Else
			pstrHREF = "<a href=""" & "detail.asp?PRODUCT_ID=" & arrProduct(0, iRec) & """>"
		End If
	End If
		
		
	'Write the Image link
	If (Len(txtImagePath) > 0) Then
		pstrImage = "  " & pstrHREF & "<img border=""0"" src=""" & txtImagePath & """ alt=""" & Server.HTMLEncode(arrProduct(1, iRec)) & """></a>" & vbcrlf
	Else
		If Len(cstrImageNotAvailablePath) > 0 Then
			pstrImage = "  " & pstrHREF & "<img border=""0"" src=""" & cstrImageNotAvailablePath & """ alt=""" & "No Image Available" & """></a>" & vbcrlf
		Else
			pstrImage = "  " & pstrHREF & "<i>No Image Available</i></a>" & vbcrlf
		End If
	End If

	'Write the text link
	pstrProductName = "  <br />" & pstrHREF & arrProduct(1, iRec) & "</a>" & vbcrlf

	'Fix if null in on sale
	If isNull(arrProduct(5, iRec)) Then arrProduct(5, iRec) = 0

	'Create the price
	If cBool(arrProduct(5, iRec)) Then 'is product on sale
		If IsNull(FormatCurrency(arrProduct(4, iRec))) Then
			pstrProductPrice = "  <br /><strike>Please contact customer service</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		Else
			pstrProductPrice = "  <br /><strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>" & vbcrlf
		End If
	Else
		pstrProductPrice = "  <br />" & FormatCurrency(arrProduct(4, iRec)) & vbcrlf
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
Dim pblnDisableOrdering:	pblnDisableOrdering = True

If cblnDisableOrderingIfAttributes Then	pblnDisableOrdering = CBool(Len(pstrAttributes) > 0)

If Not cblnAllowOrdering Then pblnDisableOrdering = False

pstrProductTemplate = "<p class=shopCenterDesc><a href='detail.asp?PRODUCT_ID={productID}' class=nav><img border=1 src='{productImage}' width=110 height=110 style='border-color:#C4C4C4'><br/>{productName}</a></p>" _
					& "<p class=shopCenterDesc><strong>{productPrice}</strong> | #{productID}<br/>" _
					& "Qty <input type=text name='QUANTITY.{productID}' title=Quantity size=3 style='font-family: Verdana, Geneva, Helvetica, sans-serif; font-size: 10px; color: #000000; background-color: #FFFFFF; border: 1px solid #777777; margin: 0; padding: 0; height: 15px; max-height: 15px; vertical-align: middle' />&nbsp;&nbsp;&nbsp;<input type=submit name=AddProduct value=ORDER border=1 class=sfImageOutline /></p>"

	If True Then
		pstrProductTemplate = Replace(pstrProductTemplate, "{productID}", arrProduct(0, iRec))
		pstrProductTemplate = Replace(pstrProductTemplate, "{productName}", arrProduct(1, iRec))
		pstrProductTemplate = Replace(pstrProductTemplate, "{productDesc}", arrProduct(11, iRec))
		pstrProductTemplate = Replace(pstrProductTemplate, "{productURL}", arrProduct(4, iRec))
		pstrProductTemplate = Replace(pstrProductTemplate, "{productImage}", txtImagePath)

		If cBool(arrProduct(5, iRec)) Then 'is product on sale
			If IsNull(FormatCurrency(arrProduct(4, iRec))) Then
				pstrProductTemplate = Replace(pstrProductTemplate, "{productPrice}", "<strike>Please contact customer service</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>")
			Else
				pstrProductTemplate = Replace(pstrProductTemplate, "{productPrice}", "<strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike>&nbsp;<font color='red'>" & FormatCurrency(arrProduct(6, iRec),2) & "</font>")
			End If
		Else
			pstrProductTemplate = Replace(pstrProductTemplate, "{productPrice}", FormatCurrency(arrProduct(4, iRec)))
		End If 

		Response.Write pstrProductTemplate
	Else
		Response.Write pstrImage
		Response.Write pstrProductName
		If Not pblnDisableOrdering Then Response.Write pstrAttributes & vbcrlf
		Response.Write pstrProductPrice & vbcrlf
		'If Not pblnDisableOrdering Then Response.Write "<br />Quantity:&nbsp;<input style='" & C_FORMDESIGN & "'  type='text' name='QUANTITY." & arrProduct(0, iRec) & "' title='Quantity' size='3' value='' onblur='return isInteger(this, true, " & Chr(34) & "Please enter an integer greater than one for the quantity" & Chr(34) & ")'>"
		'If Not pblnDisableOrdering And cblnShowIndividualAddToCartButtons Then Response.Write cstrQty_AddToCartButtonSeparator & "<input type='image' name='AddProduct' border='0' src='" & C_BTN03 & "' alt='Add To Cart'><br />"
		'If Len(cstrShowMoreDetails) > 0 Then Response.Write "<br />" & pstrHREF & cstrShowMoreDetails & "</a>"
	End If
	
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

	Response.Write "</td>" & vbcrlf
	
	If mbytCurrentColumn = cbytNumColumns Then
		If Len(mstrSpacerRow) > 0 Then Response.Write mstrSpacerRow
		Response.Write "</tr>" & vbcrlf
		mbytCurrentColumn = 0
	End If
	
	If (iRec=iVarPageSize-1) Then
		If mbytCurrentColumn > 0 Then
			For i=1 to (cbytNumColumns - mbytCurrentColumn)
				Response.Write "<td>&nbsp;</td>" & vbcrlf
			Next
		End If
		Response.Write "</table></center>" & vbcrlf
	End If
	
End Sub	'WritessSearchGrid

'***************************************************************************************************************************************************************************

Sub WritessTabularSearchResults

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
		pstrImage = "  " & pstrHREF & "<img border=""0"" src=""" & txtImagePath & """ alt=""" & Server.HTMLEncode(arrProduct(0, iRec)) & """></a>" & vbcrlf
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

	If (iRec=iVarPageSize-1) Then 
		Response.Write "</table></center>" & vbcrlf
	Else
		If cblnShowSpacerHR Then
			Response.Write "    <tr><td align='center' colspan='" & plngCellCount & "'><hr color='#6699CC' width='100%'></td></tr>" & vbcrlf
		End If
	End If
	
End Sub	'WritessTabularSearchResults

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

