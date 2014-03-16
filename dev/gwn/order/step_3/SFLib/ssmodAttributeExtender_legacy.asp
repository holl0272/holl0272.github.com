<%
'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'Custom attribute fields
	Const cstrattrURL_Field = "attrURL"
	Const cstrattrExtra_Field = "attrDisplay"
	
	'Custom attribute detail fields
	Const cstrattrdtImage_Field = "attrdtImage"
	Const cstrattrdtWeight_Field = "attrdtWeight"
	Const cstrattrdtURL_Field = "attrdtURL"
	Const cstrattrdtExtra_Field = "attrdtDisplay"
	Const cstrattrdtDefault_Field = "attrdtDefault"

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

Function adjustDetailSQL_AttributeExtender(ByVal strSQL)

Dim pstrSQL
Dim pstrFieldsToAdd

	If True Then
		adjustDetailSQL_AttributeExtender = "SELECT sfAttributes.*, sfAttributeDetail.* " _
				& "FROM sfAttributes INNER JOIN sfAttributeDetail ON sfAttributes.attrId = sfAttributeDetail.attrdtAttributeId " _
				& "WHERE attrProdId = '" & makeInputSafe(txtProdId) & "'" _
				& "ORDER BY attrDisplayOrder, AttrName, attrdtOrder"
		Exit Function
	End If
	
	pstrFieldsToAdd = " attrDisplayStyle,"

	'Attribute Fields
	If Len(cstrattrURL_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrURL_Field & ","
	If Len(cstrattrExtra_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrExtra_Field & ","
	
	'Attribute Detail Fields
	If Len(cstrattrdtImage_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrdtImage_Field & ","
	If Len(cstrattrdtWeight_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrdtWeight_Field & ","
	If Len(cstrattrdtURL_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrdtURL_Field & ","
	If Len(cstrattrdtExtra_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrdtExtra_Field & ","
	If Len(cstrattrdtDefault_Field) > 0 Then pstrFieldsToAdd = pstrFieldsToAdd & " " & cstrattrdtDefault_Field & ","

	pstrSQL = Replace(strSQL, "attrName,", "attrName," & pstrFieldsToAdd, 1, 1)
	
	'Now for the Order By clause
	pstrSQL = Replace(pstrSQL, "ORDER BY AttrName", , 1, 1)

	adjustDetailSQL_AttributeExtender = pstrSQL
			
End Function	'adjustDetailSQL_AttributeExtender

'**********************************************************************************************************

Function attributeLink(ByRef objRS, ByVal strName, ByVal blnAttribute)

Dim pstrOut
Dim pstrURL
Dim pstrExtra

	If blnAttribute Then
		If Len(cstrattrURL_Field) > 0 Then pstrURL = Trim(objRS.Fields(cstrattrURL_Field).Value & "")
		If Len(cstrattrExtra_Field) > 0 Then pstrExtra = Trim(objRS.Fields(cstrattrExtra_Field).Value & "")
		If Len(pstrExtra) > 0 Then
			pstrOut = pstrExtra
		Else
			If Len(pstrURL) > 0 Then
				pstrOut = pstrOut & "<a class=attExtAttributeAnchor href=" & pstrURL & " " & pstrExtra & ">" & strName & "</a>"
			Else
				'pstrOut = pstrOut & "<font face=" & C_FONTFACE4 & " color=""" & C_FONTCOLOR4 & """ SIZE=" & C_FONTSIZE4  &">" & strName & "</font>"
				pstrOut = pstrOut & "<span class=""attributeCategoryName"">" & strName & "</span>"
			End If
		End If
	Else
		If Len(cstrattrdtURL_Field) > 0 Then pstrURL = Trim(objRS.Fields(cstrattrdtURL_Field).Value & "")
		If Len(cstrattrdtExtra_Field) > 0 Then pstrExtra = Trim(objRS.Fields(cstrattrdtExtra_Field).Value & "")
		If Len(pstrExtra) > 0 Then
			pstrOut = pstrExtra
		Else
			If Len(pstrURL) > 0 Then
				pstrOut = pstrOut & "<a class=attExtAttributeDetailAnchor href=" & pstrURL & " " & pstrExtra & ">" & strName & "</a>"
			Else
				'pstrOut = pstrOut & "<font face=" & C_FONTFACE4 & " color=""" & C_FONTCOLOR4 & """ SIZE=" & C_FONTSIZE4  &">" & strName & "</font>"
				pstrOut = pstrOut & "<span class=""attributeCategoryName"">" & strName & "</span>"
			End If
		End If
	End If
	
	attributeLink = pstrOut

End Function	'attributeLink

'**********************************************************************************************************

Function DisplayAttributeNameSearchResults(ByVal lngCounter)

Dim pstrAttrURL
Dim pstrAttrTitle

	If Len(cstrattrURL_Field) > 0 Then pstrAttrURL = Trim(arrAtt(1, lngCounter) & "")
	
	If Len(pstrAttrURL) > 0 Then
		pstrAttrTitle = "<a href=" & pstrAttrURL & ">" & arrAtt(1, lngCounter) & "</a>"
	Else
		pstrAttrTitle = arrAtt(1, lngCounter)
	End If
	
	DisplayAttributeNameSearchResults = pstrAttrTitle

End Function	'DisplayAttributeNameSearchResults

'**********************************************************************************************************

Function DisplayAttributesSearchResults(ByVal strFormName)

Dim pbytattrDisplayStyle
Dim pstrattrName		
Dim pstrattrName_MPOh		
Dim pstrAttrTitle
Dim pstrOutput
Dim pblnChecked
Dim pstrFieldName

Dim pstrAttrDisplay
Dim pstrAttrDisplayAlt
Dim pstrAttrPrice

	pblnChecked = False
	pstrAttrTitle = arrAtt(1, iAttCounter)
	pbytattrDisplayStyle = arrAtt(3, iAttCounter)
	If Len(pbytattrDisplayStyle & "") = 0 Then pbytattrDisplayStyle = enAttrDisplay_Select
	
	pstrattrName = "attr" & icounter
	
	'Use format
	'attr# + cstrSSMPOAttributeDelimiter & prodID + cstrSSTextBasedAttributeHTMLDelimiter & attrID
	If CBool(Len(cstrSSMPOAttributeDelimiter) > 0) Then pstrattrName = pstrattrName & cstrSSMPOAttributeDelimiter & arrProduct(0, iRec)
	
	'If vDebug = 1 Then Response.Write "pbytattrDisplayStyle: " & pbytattrDisplayStyle & "<br />"
	'If vDebug = 1 Then Response.Write "pstrFieldName: " & pstrFieldName & "<br />"

	'modified to show entire price in dropdown if only one drop down present
	'this has not been tested against anything but detail.asp
	Dim pdblBasePrice
	If arrProduct(5, iRec) = "1" Then 
		pdblBasePrice = arrProduct(6, iRec)
	Else
		pdblBasePrice = arrProduct(4, iRec)
	End If 

	Select Case pbytattrDisplayStyle
		Case enAttrDisplay_Select: 'select
				pstrOutput = pstrOutput & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "'>" & vbcrlf
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
								sAmount = ""
						Select Case arrAttDetail(4, iAttDetailCounter)
							Case 1 
								sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
							Case 2 
								sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						End Select
						pstrOutput = pstrOutput & "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>" & vbcrlf
					End If
				Next
				pstrOutput = pstrOutput & "</select>" & vbcrlf
		Case enAttrDisplay_Radio:	'radio
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
								sAmount = ""
						Select Case arrAttDetail(4, iAttDetailCounter)
							Case 1 
								sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
							Case 2 
								sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						End Select
						If pblnChecked Then
							pstrOutput = pstrOutput & "<input type='radio' name='" & pstrattrName & "' value=""" & arrAttDetail(0, iAttDetailCounter) & """>" & arrAttDetail(2, iAttDetailCounter) & sAmount & "<br />" & vbcrlf
						Else
							pstrOutput = pstrOutput & "<input type='radio' name='" & pstrattrName & "' value=""" & arrAttDetail(0, iAttDetailCounter) & """ checked>" & arrAttDetail(2, iAttDetailCounter) & sAmount & "<br />" & vbcrlf
							pblnChecked = True
						End If
					End If
				Next
		Case enAttrDisplay_Text:	'text - required
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, false);" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeTitle(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf
						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
					End If
				Next
		Case enAttrDisplay_TextOpt:	'text - optional
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf

						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, true);" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeTitle(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf

						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = true;" & vbcrlf
					End If
				Next
		Case enAttrDisplay_Textarea:	'textarea - required
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, false);" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeTitle(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf

						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
					End If
				Next
		Case enAttrDisplay_TextareaOpt:	'textarea - optional
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, true);" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeTitle(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf

						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
					End If
				Next
		Case enAttrDisplay_Checkbox:	'checkbox
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<input type='checkbox' name='" & pstrFieldName & "' value='on'>" & pstrAttrTitle & strAttrPrice & "<br />" & vbcrlf
					End If
				Next
		Case enAttrDisplay_SelectShowPrice:	'select (show price)
				pstrOutput = pstrOutput & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "'>" & vbcrlf
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					Select Case arrAttDetail(4, iAttDetailCounter)
						Case 1
							pstrAttrDisplayAlt = " " & FormatCurrency( CDbl(pdblBasePrice) + CDbl(arrAttDetail(3, iAttDetailCounter))) & ""
							pstrAttrDisplay = " (Add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						Case 2
							pstrAttrDisplayAlt = " " & FormatCurrency( CDbl(pdblBasePrice) - CDbl(arrAttDetail(3, iAttDetailCounter))) & ""
							pstrAttrDisplay = " (Subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						Case Else
							pstrAttrDisplay = ""
							pstrAttrDisplayAlt = " " & FormatCurrency(pdblBasePrice) & ""
					End Select

					If irsSearchAttRecordCount = 1 And cblnssDisplayFullPriceInSingleAttributes Then
						pstrAttrPrice = pstrAttrDisplayAlt
					Else
						pstrAttrPrice = pstrAttrDisplay
					End If
	
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<option value=""" & arrAttDetail(0, iAttDetailCounter) & """>" & pstrAttrTitle & pstrAttrDisplayAlt & "</option>" & vbcrlf
					End If
				Next
				pstrOutput = pstrOutput & "</select>" & vbcrlf
		Case enAttrDisplay_SelectChangeImage:	'custom example
				pstrOutput = pstrOutput & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "' id='" & pstrattrName & "' onchange=""changeCustomImage('" & arrProduct(0, iRec) & "', this);"">" & vbcrlf
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
								sAmount = ""
						Select Case arrAttDetail(4, iAttDetailCounter)
							Case 1 
								sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
							Case 2 
								sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						End Select
						pstrOutput = pstrOutput & "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & arrProduct(0, iRec) & "'] = new Array();" & vbcrlf
					End If
				Next
				pstrOutput = pstrOutput & "</select>" & vbcrlf
		Case enAttrDisplay_Custom:	'custom example
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & arrAttDetail(0, iAttDetailCounter)
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
						pstrOutput = pstrOutput & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
						pstrOutput = pstrOutput & "<a href=""#"" onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('button_r1_c1','','images/button_r1_c1_f2.gif',1);"" ><img name=""button_r1_c1"" src=""images/button_r1_c1.gif"" width=""187"" height=""43"" border=""0"" onMouseDown=""MM_openBrWindow('/charm_directorynewmb.asp','','width=550,height=500')""></a></td>"
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "mtxtCharm = document." & strFormName & "." & pstrFieldName & ";" & vbcrlf
					End If
				Next
		Case Else: 
				pstrOutput = pstrOutput & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "'>" & vbcrlf
				For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
					If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
								sAmount = ""
						Select Case arrAttDetail(4, iAttDetailCounter)
							Case 1 
								sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
							Case 2 
								sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
						End Select
						pstrOutput = pstrOutput & "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>" & vbcrlf
					End If
				Next
				pstrOutput = pstrOutput & "</select>" & vbcrlf
	End Select
	
	DisplayAttributesSearchResults = pstrOutput

End Function	'DisplayAttributesSearchResults

'**********************************************************************************************************
%>
