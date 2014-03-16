Option Explicit

Const tvwChild = 4

Const enStatusMessage_Background = 0
Const enStatusMessage_Primary = 1
Const enStatusMessage_Primary_Error = 2

Const enTV_ID = 0
Const enTV_Parent = 1
Const enTV_Name = 2

Const cstrRootElementName = "Category"
Const cstrRootKeyName = "id"

Const cstrKeyPrefix = "id"
Const cbytDefaultRootID = "0"

Dim mobjTreeview
Dim mobjActiveTreeNode
Dim mobjActiveTreeNodeForNew

Dim mobjXMLDoc
Dim mobjActiveXMLNode

Dim maryTreeViewData(2)	'id, parent, name
Dim maryDataElement

'for edit operations
Dim mblnItemIsDirty:	mblnItemIsDirty = False
Dim mblnDataSetIsDirty:	mblnDataSetIsDirty = False

Dim mlngSelectedID

'Data specific variables


	'********************************************************************************************

	Sub ShowXMLData
		document.all("xmlData").value = mobjXMLDoc.xml
	End Sub	'ShowXMLData

	'********************************************************************************************

	Function SaveChanges()

		Call GetFormData
		
		On Error Resume Next

		Call SetTVDataArray(mobjActiveXMLNode)
		mobjActiveTreeNode.Text = maryTreeViewData(enTV_Name)
		Call MakeItemDirty(False)
		Call ShowXMLData

	End Function	'SaveChanges

	'********************************************************************************************

	Function CheckSaveChanges()

	Dim pblnResult
	Dim pblnResponse
	
		pblnResult = True
		
		If mblnItemIsDirty Then
			pblnResponse = msgbox("Do you wish to save your changes to " & mobjActiveTreeNode.Text & "?",vbYesNoCancel,"Save Changes?")
			Select Case pblnResponse
				Case vbYes	'Save Changes
					Call SaveChanges
				Case vbNo	'Abandon Changes
					Call AbandonChanges
				Case vbCancel	'return to original
					'this doesn't seem to work properly
					mobjActiveTreeNode.Selected = True
					pblnResult = False
			End Select
		End If
		
		CheckSaveChanges = pblnResult
		
	End Function	'CheckSaveChanges

	'********************************************************************************************

	Function AbandonChanges()
		Call SetFormData()
		Call MakeItemDirty(False)
	End Function	'AbandonChanges

	'********************************************************************************************

	Function setKey(byVal vntID)
		setKey = cstrKeyPrefix & vntID
	End Function	'setKey

	'********************************************************************************************

	Function getKey(byVal vntID)
		getKey = Replace(vntID, cstrKeyPrefix, "", 1, 1)
	End Function	'getKey

	'********************************************************************************************

	Sub ClearDataArray(byRef aryData)
	
	Dim i	

		For i = 0 To UBound(aryData)
			aryData(i) = ""
		Next 'i

	End Sub	'ClearDataArray

	'********************************************************************************************

	Sub SetTVDataArray(byRef objXMLNode)
	
	Dim i
	Dim e
	
		Call ClearDataArray(maryTreeViewData)
		For i = 0 To objXMLNode.childNodes.length - 1
			Set e = objXMLNode.childNodes.Item(i)
			Select Case e.nodeName
				Case "id": maryTreeViewData(enTV_ID) = e.text
				Case "parentID": maryTreeViewData(enTV_Parent) = e.text
				Case "Name": maryTreeViewData(enTV_Name) = e.text
			End Select
		Next 'i

		If Len(maryTreeViewData(enTV_Parent)) = 0 Then maryTreeViewData(enTV_Parent) = cbytDefaultRootID

	End Sub	'SetTVDataArray

	'********************************************************************************************

	Function GetXMLNodeByKey(byVal strKey)

	Dim nodeList
	Dim i
	
		Set nodeList = mobjXMLDoc.getElementsByTagName(cstrRootElementName)

		For i = 0 To nodeList.length - 1
			'If nodeList.Item(i).attributes.item(0).nodeValue = strKey Then
			If nodeList.Item(i).attributes.item(0).nodeValue = strKey Then
				Set GetXMLNodeByKey = nodeList.Item(i)
				Exit Function
			End If
		Next 'i
		
		Set nodeList = Nothing
		
	End Function	'GetXMLNodeByKey

	'********************************************************************************************

	Function RetrieveRemoteData(byVal strURL, byVal strFormData, byVal blnPostData, byVal blnRandom)
	
	Dim pobjXMLHTTP
	
	'this is here to prevent data caching problem
	Dim pstrURL
	Dim pstrAppend
	
	If blnRandom Then
		Randomize   ' Initialize random-number generator.
		pstrAppend = Int(2147483647 * Rnd)
		If Instr(1,strURL,"?") > 0 Then
			pstrURL = strURL & "&RandomTrash=" & pstrAppend
		Else
			pstrURL = strURL & "?RandomTrash=" & pstrAppend
		End If
	Else
		pstrURL = strURL
	End If
	
	'set timeouts in milliseconds
	Const resolveTimeout = 1000
	Const connectTimeout = 1000
	Const sendTimeout = 1000
	Const receiveTimeout = 10000
	
	On Error Resume Next
	
		SetStatusMessage "Retrieving remote data . . .", enStatusMessage_Background
		Set pobjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
		With pobjXMLHTTP
			If blnPostData Then
				.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
				.open "POST", "", False
				.open "POST", pstrURL, False
				.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
				.send strFormData
			Else
				.open "GET", pstrURL, False
				.send
			End If
			.ContentType = "text/xml"
			RetrieveRemoteData  = .responseText
		end with
		set pobjXMLHTTP = nothing
		SetStatusMessage "", enStatusMessage_Background

	End Function	'RetrieveRemoteData


	'********************************************************************************************

	Function LoadData

	Dim pstrRawData
	Dim pstrMessage
	Dim pblnResult
	

		pstrRawData = RetrieveRemoteData("testData/Categories.xml?Action=GetData","",False,True)

		set mobjXMLDoc = CreateObject("MSXML.DOMDocument")
		mobjXMLDoc.async = false

		SetStatusMessage "Loading data . . .", enStatusMessage_Background
		If mobjXMLDoc.loadXML(pstrRawData) Then
			pblnResult = True
		Else
			pstrMessage = "Error loading data: Error " & mobjXMLDoc.parseError.errorCode & " - " & mobjXMLDoc.parseError.reason & vbcrlf _
				    & "Line " & mobjXMLDoc.parseError.line & ", Line Position " & mobjXMLDoc.parseError.linePos & vbcrlf _
				    & "srcText: " & mobjXMLDoc.parseError.srcText

			MsgBox pstrMessage,vbOKOnly,"Error"
			pblnResult = False
		End If
		SetStatusMessage "", enStatusMessage_Background
		
		LoadData = pblnResult
		
	End Function	'LoadData

	'********************************************************************************************

	Function SetStatusMessage(byVal strMessage, byVal bytType)

	Dim pobjDivMessage

		Set pobjDivMessage = document.all("divMessage")
		Select Case bytType
			Case enStatusMessage_Background
				window.status = strMessage
			Case enStatusMessage_Primary_Error
				pobjDivMessage.innerHTML = strMessage
				pobjDivMessage.className = "messageError"
			Case enStatusMessage_Primary
				pobjDivMessage.innerHTML = strMessage
				pobjDivMessage.className = "message"
		End Select

	End Function	'SetStatusMessage

	'********************************************************************************************

	Function LoadListView
	
	Dim nodeList
	Dim i
	Dim pstrResult
	Dim pblnSuccess

	On Error Goto 0
	
		pblnSuccess = True
		Set nodeList = mobjXMLDoc.getElementsByTagName(cstrRootElementName)
		
		mobjTreeview.Nodes.Clear()
		
		If nodeList.length >= 0 Then
			ReDim maryDataElement(nodeList.Item(0).childNodes.length)
		End If

		For i = 0 To nodeList.length - 1
			Call SetTVDataArray(nodeList.Item(i))
			If Len(maryTreeViewData(enTV_Parent)) > 0 Then
				If mobjTreeview.Nodes.count = 0 Then mobjTreeview.Nodes.Add ,, setKey(cbytDefaultRootID), cstrRootElementName
				mobjTreeview.Nodes.Add setKey(maryTreeViewData(enTV_Parent)), tvwChild, setKey(maryTreeViewData(enTV_ID)), maryTreeViewData(enTV_Name)
			Else
				mobjTreeview.Nodes.Add ,, setKey(maryTreeViewData(enTV_ID)), maryTreeViewData(enTV_Name)
			End If

			If Err.number <> 0 Then
				Select Case Err.number
					Case 35601	'Element Not Found
						Err.Clear
					Case 35602	'Key Is Not Unique in Collection
						'ignore this one
						Err.Clear
					Case Else
						MsgBox "Error " & Err.number & ": " & Err.Description,vbOKOnly,"Error"
						Err.Clear
					End Select
			End If
		Next 'i

		If Not isObject(mobjActiveXMLNode) Then Set mobjActiveXMLNode = nodeList(0)

		'Expand all nodes
		For i = 1 To mobjTreeview.Nodes.Count
			mobjTreeview.Nodes.Item(i).Expanded = True
		Next 'i
	
		Set nodeList = Nothing
		LoadListView = pblnSuccess
		
	End Function	'LoadListView

	'********************************************************************************************

	Sub SetFormData()
	'Reads form contents from active XML Node to form and data array	
	
	Dim i
	Dim e
	Dim itemValue
	Dim itemName
	
		Call ClearDataArray(maryDataElement)
		For i = 0 To mobjActiveXMLNode.childNodes.length - 1
			Set e = mobjActiveXMLNode.childNodes.Item(i)
			itemName = e.nodeName
			itemValue = e.text

			maryDataElement(i) = itemValue
			document.all(itemName).value = maryDataElement(i)

		Next 'i

	End Sub	'SetFormData

	'********************************************************************************************

	Sub GetFormData()
	'Reads form contents into active XML Node and data array	

	Dim i
	Dim e
	Dim itemValue
	Dim itemName
	
		Call ClearDataArray(maryDataElement)

		For i = 0 To mobjActiveXMLNode.childNodes.length - 1
			Set e = mobjActiveXMLNode.childNodes.Item(i)

			itemName = e.nodeName
			itemValue = document.all(itemName).value
			e.text = itemValue

			maryDataElement(i) = itemValue

		Next 'i

	End Sub	'GetFormData

	'********************************************************************************************

	Sub SetFirstListItem()
	
		If mobjTreeview.Nodes.Count > 1 Then 
			mobjTreeview.Nodes.Item(1).Expanded = True
			Set mobjActiveTreeNode = mobjTreeview.Nodes.Item(2)
			Set mobjActiveTreeNodeForNew = mobjActiveTreeNode
			mobjActiveTreeNode.Selected = True
			SetStatusMessage "Setting first item active - " & getKey(mobjActiveTreeNode.Key), enStatusMessage_Background
			Set mobjActiveXMLNode = GetXMLNodeByKey(getKey(mobjActiveTreeNode.Key))
			Call SetTVDataArray(mobjActiveXMLNode)
			Call SetFormData()
			frmData.btnDelete.disabled = False
			frmData.btnCopy.disabled = False
		End If

	End Sub	'SetFirstListItem

	'********************************************************************************************

	Sub TreeView1_Click

	Dim pstrKey
	Dim pobjSourceNode

	'On Error Resume Next

		pstrKey = getKey(TreeView1.SelectedItem.key)
		Set pobjSourceNode = GetXMLNodeByKey(getKey(pstrKey))
'Exit Sub
		If Not CheckSaveChanges Then Exit Sub


		Call SetTVDataArray(mobjActiveXMLNode)
		Call SetFormData()

		
'		Set pobjSourceNode = GetListItem(mlngX,mlngY)
		Set pobjSourceNode = TreeView1.SelectedItem
		If mblnItemIsDirty Then
			If mobjActiveTreeNode.Key <> pobjSourceNode.Key Then
			End If
		End If
			
		If Not pobjSourceNode is Nothing Then
			Set mobjActiveTreeNodeForNew = pobjSourceNode
			pstrKey = pobjSourceNode.Key
			If pstrKey = "root:0" Then Exit Sub
			Set mobjActiveXMLNode = GetXMLNodeByKey(getKey(pstrKey))
			If Err.number <> 0 Then
				msgbox "Error " & Err.number & ": " & Err.Description & " (" & pstrKey & ")"
				Err.Clear
			End If
			
			Call SetTVDataArray(mobjActiveXMLNode)
			Call SetFormData()
			Set mobjActiveTreeNode = pobjSourceNode
		End If

	End Sub	'TreeView1_Click

	'********************************************************************************************

	Sub AutoUpdate()
	
		If document.frmData.chkAutoUpdate.checked Then
			msgbox "This will automatically save changes to the item as you make them." & vbcrlf _
				 & "You will lose the ability to use the reset button.",vbOKOnly,"Caution"
			frmData.btnReset.disabled = document.frmData.chkAutoUpdate.checked
			Call SaveChanges
		End If
	
	End Sub	'AutoUpdate

	'********************************************************************************************

	Sub ChangeItem()

		If document.frmData.chkAutoUpdate.checked Then
			SaveChanges()
		Else
			Call MakeItemDirty(True)
		End If
		Call MakeDataSetDirty(True)

	End Sub	'ChangeItem

	'********************************************************************************************

	Sub MakeItemDirty(byVal blnDirty)
	
		If mblnItemIsDirty and Not blnDirty Then
			'clean item
			Call SetFormData()
			mblnItemIsDirty = False
			frmData.btnReset.disabled = (Not mblnItemIsDirty) OR document.frmData.chkAutoUpdate.checked
			frmData.btnUpdateItem.disabled = Not mblnItemIsDirty
		ElseIf Not mblnItemIsDirty and blnDirty Then
			'make item dirty
			
			mblnItemIsDirty = True
			frmData.btnReset.disabled = (Not mblnItemIsDirty) OR document.frmData.chkAutoUpdate.checked
			frmData.btnUpdateItem.disabled = Not mblnItemIsDirty
		End If

	End Sub	'MakeItemDirty

	'********************************************************************************************

	Sub MakeDataSetDirty(byVal blnDirty)
	
		If mblnDataSetIsDirty and Not blnDirty Then
			'clean item
			
			mblnDataSetIsDirty = False
			frmData.btnUpdateDataset.disabled = Not mblnDataSetIsDirty
		ElseIf Not mblnDataSetIsDirty and blnDirty Then
			'make item dirty
			
			mblnDataSetIsDirty = True
			frmData.btnUpdateDataset.disabled = Not mblnDataSetIsDirty
		End If

	End Sub	'MakeDataSetDirty

	'********************************************************************************************	
	sub window_onLoad

		'initialize variables
		Set mobjTreeview = Treeview1

		SetStatusMessage "Loading data . . .", enStatusMessage_Primary
		If LoadData Then
			If LoadListView Then
				Call SetFirstListItem
				SetStatusMessage "", enStatusMessage_Primary
			Else
				pblnResult = False
				SetStatusMessage "<h4><font color=red>Error loading List</font></h4>"
			End If
			SetStatusMessage "", enStatusMessage_Primary
		Else
			SetStatusMessage "Error loading data", enStatusMessage_Primary_Error
		End If


	End Sub	'window_onLoad

	'********************************************************************************************

