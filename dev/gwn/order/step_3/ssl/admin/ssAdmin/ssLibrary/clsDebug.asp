<%
Class clsDebug

	Dim mb_Enabled
	Dim md_RequestTime
	Dim md_FinishTime
	Dim mo_Storage

	Public Default Property Get Enabled()
		Enabled = mb_Enabled
	End Property
	Public Property Let Enabled(bNewValue)
		mb_Enabled = bNewValue
	End Property

	Private Sub Class_Initialize()
		md_RequestTime = Now()
		Set mo_Storage = Server.CreateObject("Scripting.Dictionary")
		mb_Enabled = True
	End Sub

	Public Sub PrintLater(label, output)
		If Enabled Then
			'save output to internal dictionary
			Call mo_Storage.Add(label, output)
		End If
	End Sub

	Public Sub Print(label, output)
		Response.Write "<h3>" & label & ": " & output & "</h3><br>" & vbcrlf
	End Sub

	Public Sub [End]()
		md_FinishTime = Now()
		If Enabled Then
			Call PrintSummaryInfo()
			Call PrintCollection ("VARIABLE STORAGE", mo_Storage)
			Call PrintCollection ("QUERYSTRING COLLECTION", Request.QueryString())
			Call PrintCollection("FORM COLLECTION", Request.Form())
			Call PrintCollection ("COOKIES COLLECTION", Request.Cookies())
			Call PrintCollection ("SERVER VARIABLES COLLECTION", Request.ServerVariables())
		End If
	End Sub

	Public Sub PrintStorage()
		If Enabled Then Call PrintCollection ("VARIABLE STORAGE", mo_Storage)
	End Sub

	Public Sub PrintQuerystring()
		If Enabled Then Call PrintCollection ("QUERYSTRING COLLECTION", Request.QueryString())
	End Sub

	Public Sub PrintForm()
		If Enabled Then Call PrintCollection("FORM COLLECTION", Request.Form())
	End Sub

	Public Sub PrintCookies()
		If Enabled Then Call PrintCollection ("COOKIES COLLECTION", Request.Cookies())
	End Sub

	Public Sub PrintServerVariables()
		If Enabled Then Call  PrintCollection ("SERVER VARIABLES COLLECTION", Request.ServerVariables())
	End Sub

	Public Sub PrintSummaryInfo()
		With Response
			.Write("<hr>")
			.Write("<b>SUMMARY INFO</b></br>")
			.Write("Time of Request=" & _
				md_RequestTime) & "<br>"
			.Write("Elapsed Time=" & DateDiff("s", md_RequestTime, md_FinishTime) & " seconds<br>")
			.Write("Request Type=" & _
				Request.ServerVariables _
				("REQUEST_METHOD") & "<br>")
			.Write("Status Code=" & Response.Status _
				& "<br>")
		End With 
	End Sub

	Private Sub PrintCollection(Byval Name, ByVal Collection)
	Dim vItem
	
	Response.Write("<br><b>" & Name & "</b><br>")
		For Each vItem In Collection
			Response.Write(vItem & "=" & Collection(vItem) & "<br>")
		Next
	End Sub

	Public Sub WriteToFile(blnNewWindow, strFile, blnOverwrite)
	
	Const cstrDebugLog = "cgi-bin\debug.txt"
	Dim fso
	Dim txtFile
	Dim pstrFilePath
	Dim pstrURL
	Dim vItem
	
		On Error Resume Next
		pstrFilePath = "D:\root\crudmaster\sandshot.net\db\debug.txt"
'		pstrFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & cstrDebugLog
		pstrURL = Request.ServerVariables("SERVER_NAME") & "/" & cstrDebugLog
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		If blnOverwrite Then
			Set txtFile = fso.OpenTextFile(strFile,2,True)
		Else
			Set txtFile = fso.OpenTextFile(strFile,8,True)
		End If
		If Err.number <> 0 Then
			debugprint Err.number,Err.Description
		End If
		For Each vItem In mo_Storage
			txtFile.WriteLine vItem & "=" & mo_Storage(vItem)
		Next
		txtFile.Close
		
		Set txtFile = Nothing
		Set fso = Nothing
		
		If blnNewWindow Then Response.Write "<script>window.open('http://" & pstrURL & "');</script>"
	
	End Sub	'WriteToFile

	Private Sub Class_Terminate()
		Set mo_Storage = Nothing
	End Sub
End Class

Dim debug

Set debug = New clsDebug
%>
