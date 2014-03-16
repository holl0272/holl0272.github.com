<%
'This file centralizes extra form field data collected on orderform.asp
'
' Fields
'	- CatalogEdition - a required dropdown
'	- CatalogID - an optional text field
'	- HowHear - a required dropdown
'
' This file must be included in verify.asp and confirm.asp


'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Dim maryCustomValues(1)	'change the number to the number of extra fields

'maryCustomValues(1) = Array("","","")	'First entry is the field name on the form
										'Second entry is the database field name
										'Third entry is to be left ""
										'Fourth entry is email text name, leave blank to not include
									
maryCustomValues(1) = Array("projectedDeliveryDate","projectedDeliveryDate","","")
'maryCustomValues(2) = Array("question2","","","Check who changes your filters out?")
									
'/
'/////////////////////////////////////////////////
		
	Call LoadCustomValues

	'***********************************************************************************************

	Sub LoadCustomValues
	
	Dim plngCounter
	
		For plngCounter = 1 to UBound(maryCustomValues)
			maryCustomValues(plngCounter)(2) = Request.Form(maryCustomValues(plngCounter)(0))
			'Response.Write maryCustomValues(plngCounter)(0) & " = " & maryCustomValues(plngCounter)(2) & "<br />"
		Next 'plngCounter

	End Sub	'LoadCustomValues

	'**********************************************************************************************************************************

	Sub SaveCustomFormFields(objRS)

	Dim plngCounter
	Dim pstrFieldName
	
		For plngCounter = 1 to UBound(maryCustomValues)
			pstrFieldName = maryCustomValues(plngCounter)(1)
			If Len(pstrFieldName) = 0 Then pstrFieldName = maryCustomValues(plngCounter)(0)
			If Len(maryCustomValues(plngCounter)(2)) > 0 Then
				objRS.Fields(pstrFieldName).Value = maryCustomValues(plngCounter)(2)
			Else
				objRS.Fields(pstrFieldName).Value = Null
			End If
		Next 'plngCounter

	End Sub	'SaveCustomFormFields

	'**********************************************************************************************************************************

	Sub SaveCustomFormFields_SP(byVal lngOrderID)

	Dim plngCounter
	Dim pobjCmd
	Dim pstrFieldName
	Dim pstrSQL
	
		On Error Resume Next
	
		If Len(CStr(lngOrderID)) = 0 Or CStr(lngOrderID) = "-1" Then Exit Sub
		
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			
			For plngCounter = 1 to UBound(maryCustomValues)
				pstrFieldName = maryCustomValues(plngCounter)(1)
				If Len(pstrFieldName) = 0 Then pstrFieldName = maryCustomValues(plngCounter)(0)
				If Len(maryCustomValues(plngCounter)(2)) > 0 Then
					If Len(pstrSQL) = 0 Then
						pstrSQL = "Update sfOrders Set " & pstrFieldName & "=?"
					Else
						pstrSQL = pstrSQL & ", " & pstrFieldName & "=?"
					End If
					.Parameters.Append .CreateParameter(pstrFieldName, adWChar, adParamInput, Len(maryCustomValues(plngCounter)(2)), maryCustomValues(plngCounter)(2))
				End If
			Next 'plngCounter

			If Len(pstrSQL) > 0 Then
			
				pstrSQL = pstrSQL & " Where orderID=?"
				.Commandtext = pstrSQL
				
				Set .ActiveConnection = cnn
			
				.Parameters.Append .CreateParameter("orderID", adInteger, adParamInput, 4, lngOrderID)

				If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, plngID, "SaveCustomFormFields_SP")
				.Execute , , adExecuteNoRecords
			End If
			
		End With	'pobjCmd
		closeobj(pobjCmd)
		
		If Err.number <> 0 Then Err.Clear

	End Sub	'SaveCustomFormFields_SP

	'**********************************************************************************************************************************

	Sub GetCustomFormFields(objRS)

	Dim plngCounter
	Dim pstrFieldName
	
		For plngCounter = 1 to UBound(maryCustomValues)
			pstrFieldName = maryCustomValues(plngCounter)(1)
			If Len(pstrFieldName) = 0 Then pstrFieldName = maryCustomValues(plngCounter)(0)
			maryCustomValues(plngCounter)(2) = Trim(objRS.Fields(pstrFieldName).Value & "")
		Next 'plngCounter

	End Sub	'GetCustomFormFields

	'**********************************************************************************************************************************

	Sub WriteCustomHiddenFormFields

	Dim plngCounter
	
		For plngCounter = 1 to UBound(maryCustomValues)
			Response.Write HTMLHiddenField(maryCustomValues(plngCounter)(0), maryCustomValues(plngCounter)(2))
		Next 'plngCounter

	End Sub	'WriteCustomHiddenFormFields

	'**********************************************************************************************************************************

	Function CustomFormField(strFieldName)

	Dim plngCounter
	
		For plngCounter = 1 to UBound(maryCustomValues)
			If strFieldName = maryCustomValues(plngCounter)(0) Then
				CustomFormField = maryCustomValues(plngCounter)(2)
				Exit For
			End If
		Next 'plngCounter

	End Function	'CustomFormField

	'**********************************************************************************************************************************

	Function WriteCustomFormValuesToEmail

	Dim pstrTemp
	Dim plngCounter
	
		For plngCounter = 1 to UBound(maryCustomValues)
			If Len(maryCustomValues(plngCounter)(3)) > 0 Then pstrTemp = pstrTemp & maryCustomValues(plngCounter)(3) & ": " & maryCustomValues(plngCounter)(2) & vbcrlf
		Next 'plngCounter

		WriteCustomFormValuesToEmail = pstrTemp
		
	End Function	'WriteCustomFormValuesToEmail

	'**********************************************************************************************************************************

	Function WriteCustomFormValuesToHTML
		WriteCustomFormValuesToHTML = Replace(WriteCustomFormValuesToEmail,vbcrlf,"<br />")
	End Function	'WriteCustomFormValuesToEmail

	'**********************************************************************************************************************************

	Function HTMLHiddenField(strFieldName, strFieldValue)
		HTMLHiddenField = "<input type=hidden id=" & chr(34) & Server.HTMLEncode(strFieldName) & chr(34) _
						& " name=" & chr(34) & Server.HTMLEncode(strFieldName) & chr(34) _
						& " value=" & chr(34) & Server.HTMLEncode(strFieldValue) & chr(34) & ">"
	End Function	'HTMLHiddenField
%>