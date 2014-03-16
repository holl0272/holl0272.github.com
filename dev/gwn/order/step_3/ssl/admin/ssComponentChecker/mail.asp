<%@Language="VBScript"%>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Advanced Mailer Component Test
' © 2002 PensaWorks, inc.
' For more information, please visit http://www.pensaworks.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
on error resume next
response.buffer = true
server.scripttimeout = 60
lastUpdate = "8/30/2003"
thisPage = mid(request.servervariables("SCRIPT_NAME"), instrrev(request.servervariables("SCRIPT_NAME"), "/") + 1)
searchURL = "http://www.pensaworks.com/support/default.asp?a=5&n=1&NR=25&strSearchDisplay=show&searchtype=any&searchin=all&ST=#SearchText#"

toEmail = request("toEmail")
toName = request("toName")
fromEmail = request("fromEmail")
fromName = request("fromName")
if (request("mailerPath") <> "") then mailerPath = request("mailerPath") else mailerPath = "mail." & replace(request.servervariables("SERVER_NAME"), "www.", "", 1, -1, 1)
if (request("mailerPort") <> "") then mailerPort = request("mailerPort") else mailerPort = "25"

if (request("emlASPMail") = "Y") then emlASPMail = true else emlASPMail = false
if (request("emlASPEmail") = "Y") then emlASPEmail = true else emlASPEmail = false
if (request("emlASPQMail") = "Y") then emlASPQMail = true else emlASPQMail = false
if (request("emlCDONTS") = "Y") then emlCDONTS = true else emlCDONTS = false
if (request("emlJMail") = "Y") then emlJMail = true else emlJMail = false
if (request("emlSASmtpMail") = "Y") then emlSASmtpMail = true else emlSASmtpMail = false
if (request("emlIPWorks") = "Y") then emlIPWorks = true else emlIPWorks = false
if (request("emlCDOSYS") = "Y") then emlCDOSYS = true else emlCDOSYS = false

emlSubject = "Advanced Mailer Component Test - #ComName#"
emlBody = emlBody & "Receipt of this email is proof that the component is installed and functions properly with the settings below. If you believe you have reached this email in error, then please disregard it. "
emlBody = emlBody & VbCrLf
emlBody = emlBody & VbCrLf & "* Mailer Component: #ComName#"
emlBody = emlBody & VbCrLf & "* To Name: " & toName
emlBody = emlBody & VbCrLf & "* To Email: " & toEmail
emlBody = emlBody & VbCrLf & "* From Name: " & fromName
emlBody = emlBody & VbCrLf & "* From Email: " & fromEmail
emlBody = emlBody & VbCrLf & "* Mailer Path: " & mailerPath
emlBody = emlBody & VbCrLf & "* Mailer Port: " & mailerPort
emlBody = emlBody & VbCrLf & VbCrLf & "Sent from the Advanced Mailer Component Test on " & Now & "."
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Advanced Mailer Component Test - http://www.pensaworks.com</title>
</head>
<body topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<table border="0" width="100%" cellspacing="0" cellpadding="3"><tr><td width="100%" bgcolor="#000080"><div align="center"><font size="2" color="mintcream"><b><font face="Arial, Helvetica, sans-serif" size="4">Advanced Mailer Component Test</font></b></font></div></td></tr></table>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Sometimes a server may accept the CreateObject code showing that a mailer component is installed. This does not necesesarily mean it is setup and functioning properly. To ensure which mailer programs are functioning and configured properly, please fill out the form below, choose the mailer programs to test, and hit &quot;Run The Test&quot;. If the email is never received, then the mailer component may not be configured properly. Contact your host in that case with any errors printed out. Please <a href="http://www.pensaworks.com/contact.asp" target="_blank">send us</a> any feedback, bugs, or requests you may have. This script was last updated <%=lastUpdate%>.</font> 
<table border="0" align="center" width="75%" cellpadding="4"><form name="Subscribe" method="post" action="http://www.pensaworks.com/mailinglist.asp"><tr><td bgcolor="#CCCCCC"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Mailer Component Test Mailinglist</b></font><br><font size="1" face="Verdana, Arial, Helvetica, sans-serif">We are constantly making changes and updating the component tests to add new features and more components to the list. Join the mailinglist and be the first to know when we release an update! View our <a href="http://www.pensaworks.com/privacy.asp" target="_blank">Privacy Policy</a>.</font></div><div align="center"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Name:</font></b> <input type="text" name="Name" size="10">&nbsp;&nbsp;&nbsp; <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Email Address:</font></b> <input type="text" name="Email" size="25"><input type="hidden" name="L" value="5"><input type="hidden" name="Action" value="Subscribe"><input type="hidden" name="a" value="s"><input type="submit" name="Subscribe" value="Subscribe"></div></td></tr></form></table>
<form name="Test" method="post" action="<%=thisPage%>">
<table border="0" align="center">
<tr><td colspan="3">
	<table border="0" align="center">
	<tr><td><input type="checkbox" name="emlCDONTS" value="Y"<% if (emlCDONTS) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">CDONTS</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(1)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlASPMail" value="Y"<% if (emlASPMail) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">ASPMail</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(2)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlASPQMail" value="Y"<% if (emlASPQMail) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">ASPQMail</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(3)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlJMail" value="Y"<% if (emlJMail) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">JMail</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(4)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlSASmtpmail" value="Y"<% if (emlSASmtpmail) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">SA-SmtpMail</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(5)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlASPEmail" value="Y"<% if (emlASPEmail) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">ASPEmail</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(6)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlIPWorks" value="Y"<% if (emlIPWorks) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">IPWorks</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(7)%></font></td></tr>
	<tr><td><input type="checkbox" name="emlCDOSYS" value="Y"<% if (emlCDOSYS) then response.write " CHECKED" %>></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">CDOSYS</font></td><td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=isInstalled(8)%></font></td></tr>
	</table>
	<p align="center"><font size="2" color="#0000FF" face="Verdana, Arial, Helvetica, sans-serif">Make sure you use a valid To Email Address or you will never receive any<br>emails from this portion of the test. All fields are required.</font></p>
</td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">To Email Address: </font></b><font size="1">(ex: John@domain.com)</font></font></td><td><input name="toEmail" type="text" id="toEmail" value="<%=toEmail%>" size="30"></td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">To Name: </font></b><font size="1">(ex: John Doe)</font></font></td><td><input name="toName" type="text" id="toName" value="<%=toName%>" size="20"></td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">From Email Address: </font></b><font size="1">(ex: John@domain.com)</font></font></td><td><input name="fromEmail" type="text" id="fromEmail" value="<%=fromEmail%>" size="30"></td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">From Name: </font></b><font size="1">(ex: John Doe)</font></font></td><td><input type="text" name="fromName" size="20" value="<%=fromName%>"></td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Mailer Path:</font></b><font size="1">(ex: mail.yourdomain.com)</font></font></td><td><input type="text" name="mailerPath" size="30" value="<%=mailerPath%>"></td></tr>
	<tr><td><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Mailer Port: </font></b><font size="1">(25 by default, for JMail &amp; ASPEmail)</font><b></b></font></td><td><input type="text" name="mailerPort" size="5" value="<%=mailerPort%>"></td></tr>
	<tr><td colspan="3"><div align="center"><input type="hidden" name="action" value="go"><input type="submit" name="Submit" value="Run the Test"><input type="reset" name="Submit" value="Reset Form"></div></td></tr>
</table>
</form>
<% if request("action") = "go" then %>
<div align="center"><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Running Tests now. Any errors encountered will be printed out to the browser.</strong></font><br>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
<%
	if emlCDONTS then call runTest(1, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlASPMail then call runTest(2, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlASPQMail then doIt = runTest(3, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlJMail then call runTest(4, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlSASmtpMail then call runTest(5, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlASPEmail then call runTest(6, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlIPWorks then call runTest(7, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
	if emlCDOSYS then call runTest(8, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL) : err = 0
%><br><font color="#FF0000"><strong>Finished running the mailer test(s). Please check your email for the confirmation emails.</strong></font></font></div>
<% end if %>
<table border="0" width="98%" cellpadding="2" align="center"><tr><td width="200%"><hr width="90%"></td></tr><tr><td width="25%" valign="bottom"><p align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>© 2001, <a href="http://www.pensaworks.com" target="_blank">PensaWorks, inc.</a><br>All Rights Reserved.</b></font></td></tr></table>
</body></html>
<%
'======================================
' Function that test for a component by trying to create the object
function isObjInstalled(strClassString)
	on error resume next
	err = 0 : isObjInstalled = false
		set xTestObj = Server.CreateObject(strClassString)
			if (Err = 0) then isObjInstalled = true else isObjInstalled = false
		set xTestObj = nothing
	err = 0 : err.clear()
end function

'======================================
' Function to print out whether component is installed or not
function isInstalled(com)
	doIt = getCom(com, comName, createStr)
	if (isObjInstalled(createStr)) then isInstalled = "<b><font color=""#00CC00"">Installed</font></b>" else 	isInstalled = "<b><font color=""#FF0000"">Not Installed</font></b>"
end function

'======================================
' Function that passes back the component name and create string
function getCom(com, comName, createStr)
	select case com
		case 2
			createStr = "SMTPsvg.Mailer" : comName = "ASPMail"
		case 3
			createStr = "SMTPsvg.Mailer" : comName = "ASPQMail"
		case 4
			createStr = "Jmail.smtpmail" : comName = "JMail"
		case 5
			createStr = "SoftArtisans.SMTPMail" : comName = "SA-SMTPMail"
		case 6
			createStr = "Persits.MailSender" : comName = "ASPEMail"
		case 7
			createStr = "IPWorksASP.SMTP" : comName = "IPWorks"
		case 8
			createStr = "CDO.Message" : comName = "CDOSYS"
		case else
			createStr = "CDONTS.NewMail" : comName = "CDONTS"
		end select
end function

'======================================
' Function that runs the actual test on the components
function runTest(com, emlSubject, emlBody, FromName, FromEmail, ToName, ToEmail, MailerPath, MailerPort, searchURL)
	doIt = getCom(com, comName, createStr)
	Subject = replace(emlSubject, "#ComName#", comName)
	Message = replace(emlBody, "#ComName#", comName)
	select case com
		case 2
			doIt = ASPMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 3
			doIt = ASPQMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 4
			doIt = JMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 5
			doIt = SASmtpMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 6
			doIt = ASPEMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 7
			doIt = IPWorks_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case 8
			doIt = CDOSYS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		case else
			doIt = CDONTS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
		end select
		
		if (errText <> "") then response.write "<br>" & errText
		response.flush()
end function

'======================================
' Sends an email with CDONTS and passes back fail/success
function CDONTS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("CDONTS.NewMail") 
	if err.number <> 0 then
		errText = displayError("CDONTS", searchURL, err.Number, err.Source, err.Description)
		CDONTS_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.To = ToEmail
	Mailer.From = FromEmail
	Mailer.Subject = Subject
	Mailer.Body = Message
	Mailer.MailFormat = 1
	Mailer.BodyFormat = 1
	Mailer.Send
	if err.number <> 0 then
		errText = displayError("CDONTS", searchURL, err.Number, err.Source, err.Description)
		CDONTS_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	CDONTS_Mailer = true
end function

'======================================
' Sends an email with ASPMail and passes back fail/success
function ASPMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	if err.number <> 0 then
		errText = displayError("ASPMail", searchURL, err.Number, err.Source, err.Description)
		ASPMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.RemoteHost = MailerPath
	Mailer.ContentType = "text/plain"
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.AddRecipient ToName, ToEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.SendMail
	if err.number <> 0 then
		errText = displayError("ASPMail", searchURL, err.Number, err.Source, err.Description)
		ASPMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	ASPMail_Mailer = true
end function

'======================================
' Sends an email with ASPQMail and passes back fail/success
function ASPQMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	if err.number <> 0 then
		errText = displayError("ASPQMail", searchURL, err.Number, err.Source, err.Description)
		ASPQMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.RemoteHost = MailerPath
	Mailer.ContentType = "test/plain"
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.AddRecipient ToName, ToEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.QMessage = true
	Mailer.SendMail
	if err.number <> 0 then
		errText = displayError("ASPQMail", searchURL, err.Number, err.Source, err.Description)
		ASPQMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	ASPQMail_Mailer = true
end function

'======================================
' Sends an email with SASmtpMail and passes back fail/success
function SASmtpMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail")
	if err.number <> 0 then
		errText = displayError("SASmtpMail", searchURL, err.Number, err.Source, err.Description)
		SASmtpMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.RemoteHost = MailerPath
	Mailer.contenttype = "text/plain"
	Mailer.AddRecipient ToName, ToEmail
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.SendMail
	if err.number <> 0 then
		errText = displayError("SASmtpMail", searchURL, err.Number, err.Source, err.Description)
		SASmtpMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	SASmtpMail_Mailer = true
end function

'======================================
' Sends an email with JMail and passes back fail/success
function JMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("JMail.SMTPMail") 
	if err.number <> 0 then
		errText = displayError("JMail", searchURL, err.Number, err.Source, err.Description)
		JMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.ServerAddress = MailerPath & ":" & MailerPort
	Mailer.contenttype = "text/plain"
	Mailer.AddRecipient ToName & " <" & ToEmail & ">"
	Mailer.Sender = FromName & " <" & FromEmail & ">"
	Mailer.Subject = Subject
	Mailer.Body = Message
	Mailer.Execute
	if err.number <> 0 then
		errText = displayError("JMail", searchURL, err.Number, err.Source, err.Description)
		JMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	JMail_Mailer = true
end function

'======================================
' Sends an email with IPWorks and passes back fail/success
function IPWorks_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("IPWorksASP.SMTP") 
	if err.number <> 0 then
		errText = displayError("IPWorks", searchURL, err.Number, err.Source, err.Description)
		IPWorks_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.MailServer = MailerPath
	Mailer.MailPort = MailerPort
'	Mailer.contenttype = "text/plain"
	Mailer.SendTo = ToName & " <" & ToEmail & ">"
	Mailer.From = FromName & " <" & FromEmail & ">"
	Mailer.Subject = Subject
	Mailer.MessageText = Message
	Mailer.Send
	if err.number <> 0 then
		errText = displayError("IPWorks", searchURL, err.Number, err.Source, err.Description)
		IPWorks_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	IPWorks_Mailer = true
end function

'======================================
' Sends an email with ASPEmail and passes back fail/success
function ASPEmail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("Persits.MailSender")  
	if err.number <> 0 then
		errText = displayError("ASPEMail", searchURL, err.Number, err.Source, err.Description)
		ASPEMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Mailer.Host = MailerPath
	Mailer.Port = MailerPort
	Mailer.From = FromEmail
	Mailer.FromName = FromName
	Mailer.AddAddress ToEmail, ToName
	Mailer.Subject = Subject
	Mailer.Body = Message
	Mailer.Send
	if err.number <> 0 then
		errText = displayError("ASPEMail", searchURL, err.Number, err.Source, err.Description)
		ASPEMail_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	Set Mailer = Nothing
	ASPEmail_Mailer = true
end function


'======================================
' Sends an email with CDOSYS and passes back fail/success
function CDOSYS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath, MailerPort, errText, searchURL)
	on error resume next
	dim Mailer
	set Mailer = server.createobject("CDO.Message") 
	if err.number <> 0 then
		errText = displayError("CDOSYS", searchURL, err.Number, err.Source, err.Description)
		CDOSYS_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	
	Mailer.From = FromName & " <" & FromEmail & ">"
	Mailer.To = ToName & " <" & ToEmail & ">"
	Mailer.TextBody = Message
	Mailer.Subject = Subject
	with Mailer.Configuration
		.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailerPath
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MailerPort
		.Fields.Update
	end with
	Mailer.Send
	if err.number <> 0 then
		errText = displayError("CDOSYS", searchURL, err.Number, err.Source, err.Description)
		CDOSYS_Mailer = false : set Mailer = nothing : err.clear() : err = 0
		exit function
	end if
	set Mailer = Nothing
	CDOSYS_Mailer = true
end function


function displayError(comName, searchURL, errNumber, errSource, errDescription)
on error resume next
	if instr(errDescription, "Server.CreateObject Failed~Invalid ProgID") <> 0 then
		searchText = errNumber & "+" & server.urlencode("Server.CreateObject Failed Invalid ProgID")
	else
		searchText = errNumber & "+" & server.urlencode(errDescription)
	end if
	displayError = "<b>" & comName & " Error:</b><br>Error Number: " & errNumber & "<br>Error Source: " & errSource & "<br>Error Description: " & errDescription & "<br><a href=" & replace(searchURL, "#SearchText#", searchText) & ">More Info</a><br>"
end function
%>