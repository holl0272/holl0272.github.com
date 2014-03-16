<%@ LANGUAGE = VBScript %>
<%
Option Explicit

Function incrementCounter()
	i = i + 1
	incrementCounter = i
End Function
Const enCategory = 0
Const enTitle = 1
Const enObject = 2
Const enVersion = 3
Const enURL = 4

Dim i
Dim mstrCategory
Dim maryComponents(9)
Dim mobjTestObject

i = -1
maryComponents(0) = Array("Category", "Title", "Object", "Version", "URL")

maryComponents(incrementCounter) = Array("StoreFront System Components", "ADO", "ADODB.Connection", "yes", "http://www.microsoft.com/data/download.htm")
maryComponents(incrementCounter) = Array("StoreFront System Components", "File System Object", "Scripting.FileSystemObject", "", "")
maryComponents(incrementCounter) = Array("StoreFront System Components", "Scripting Dictionary", "Scripting.Dictionary", "", "")
maryComponents(incrementCounter) = Array("StoreFront System Components", "Credit Card Encryption - pre 50.5 Release", "SFServer.CCEncrypt", "", "http://www.storefront.net")
maryComponents(incrementCounter) = Array("StoreFront System Components", "Credit Card Encryption - 50.5 Release", "SFServer505.CCEncrypt", "", "http://www.storefront.net")

maryComponents(incrementCounter) = Array("Email Components", _
										 "Simple Mail", _
										 "SimpleMail.smtp.1", _
										 "", _
										 "http://www.simplemail.adiscon.com/en/")
maryComponents(incrementCounter) = Array("Email Components", _
										 "Simple Mail 2.0", _
										 "SimpleMail.smtp", _
										 "", _
										 "http://www.simplemail.adiscon.com/en/")
maryComponents(incrementCounter) = Array("Email Components", _
										 "Simple Mail 3.1", _
										 "ADISCON.SimpleMail.1", _
										 "", _
										 "http://www.simplemail.adiscon.com/en/")

maryComponents(incrementCounter) = Array("Payment Processor Components", _
										 "SecurePay", _
										 "SPCOM.clsSecureSend", _
										 "", _
										 "http://www.securepay.com/")

maryComponents(incrementCounter) = Array("Payment Processor Components", _
										 "PayPal Website Payments", _
										 "com.paypal.sdk.COMNetInterop.COMUtil", _
										 "", _
										 "http://www.paypal.com/")
%>

<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" ID="Table1">
	<TR>
		<TD VALIGN="TOP">
			<B>ServerName: </B><%=Request.ServerVariables("SERVER_NAME")%><BR>
			<B>Server Type: </B><%=Request.ServerVariables("Server_software")%><BR>
			<B>ServerProtocol: </B><%=Request.ServerVariables("SERVER_PROTOCOL")%><BR>
			<B>PathInfo: </B><% = Request.ServerVariables("PATH_INFO")%><BR>
			<B>PathTranslated: </B><%=Request.ServerVariables("PATH_TRANSLATED")%><BR>
			<B>Shared Hosting: </B>This website is site # <%=Request.ServerVariables("INSTANCE_ID")%> on the server<BR>
		</TD>
		<TD VALIGN="TOP">
			<FONT SIZE="3"><B>Script Engine</B><BR>
			<B>Type: </B><% = ScriptEngine%><BR>
			<B>Version: </B><%=ScriptEngineMajorVersion()%>.<%=ScriptEngineMinorVersion()%><BR>
			<B>Build: </B><%=ScriptEngineBuildVersion()%><BR>
		</TD>
		</TR>
</TABLE>
<%
For i = 0 To UBound(maryComponents)
	If mstrCategory <> maryComponents(i)(enCategory) Then
		mstrCategory = maryComponents(i)(enCategory)
		Response.Write "<hr /><h4>" & mstrCategory & "</h4>"
	End If
	
	On Error Resume Next
	Set mobjTestObject = Server.createobject(maryComponents(i)(enObject))
	If Err.Number = 0 Then
		if Len(maryComponents(i)(enVersion)) = 0 then
			Response.Write "<font color=green>" & maryComponents(i)(enTitle) & " is installed</font><br />"
		else
			Response.Write "<font color=green>" & maryComponents(i)(enTitle) & " Version: " & mobjTestObject.VERSION & " is installed</font><br />"
		end if		
   	Else
		Response.Write "<font color=red>" & maryComponents(i)(enTitle) & " is <b>not</b> installed. (" & maryComponents(i)(enObject) & ")</font><br />"
		If Len(maryComponents(i)(enURL)) > 0 Then Response.write "-- Vendor URL: <A HREF=" & chr(34) & maryComponents(i)(enURL) & chr(34) & "target=" & chr(34) & "_new" & chr(34) & ">" & maryComponents(i)(enURL) & "</A><br>"
	end if
	Set mobjTestObject = Nothing
Next

'Now test permissions
Dim fso
Dim txtFile
Dim filePath
Dim BaseFilePath
Dim maryPathsToCheck(2)

maryPathsToCheck(0) = Array("fpdb folder", "fpdb\", "testfile.txt")
maryPathsToCheck(1) = Array("fpdb folder", "ssl\admin\ssadmin\", "testfile.txt")
maryPathsToCheck(2) = Array("fpdb folder", "ssl\admin\ssadmin\ssExportedFiles\", "testfile.txt")

	BaseFilePath = Replace(Request.ServerVariables("PATH_TRANSLATED"), "ssl\admin\ssComponentChecker\ssComponentChecker.asp", "")

	Response.Write "<h4>Checking file permissions</h4>"
	Response.Write "Base File Path: " & BaseFilePath & "<br>"
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	For i = 0 To UBound(maryPathsToCheck)
		filePath = maryPathsToCheck(i)(1) & maryPathsToCheck(i)(2)
		Set txtFile = fso.OpenTextFile(BaseFilePath & filePath,2,True)
		txtFile.WriteLine "test write"
		txtFile.Close
		Set txtFile = Nothing
		If Err.Number = 0 Then
			Response.Write "<font color=green>Successfully wrote test file to " & filePath & ".</font><br />"
			Set txtFile = fso.OpenTextFile(BaseFilePath & filePath,8,True)
			txtFile.WriteLine "test write"
			txtFile.Close
			Set txtFile = Nothing
			If Err.Number = 0 Then
				Response.Write "<font color=green>Successfully updated test file at " & filePath & ".</font><br />"
   			Else
				Response.Write "<font color=red>Modify permissions are not enabled on " & filePath & "</font><br />"
				Err.Clear
			end if
   		Else
			Response.Write "<font color=red>Write permissions are not enabled on " & filePath & "</font><br />"
			Err.Clear
		end if
		
		fso.DeleteFile(BaseFilePath & filePath)
		If Err.Number = 0 Then
			Response.Write "<font color=green>Successfully deleted test file at " & filePath & ".</font><br />"
   		Else
			Response.Write "<font color=red>Delete permissions are not enabled on " & filePath & "</font><br />"
			Err.Clear
		end if
		
	Next 'i
	Set fso = Nothing
%>