<%Option Explicit 

sub debugprint(sField1,sField2)
Response.Write "<H3>" & sField1 & ": " & sField2 & "</H3><BR>"
end sub

Const dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"

dim conn,strConn,strDBpath,sql
dim strTableName,strFieldName
Dim strMessage
dim mstrAction
dim mblnError
dim mblnValidConnection
dim mstrFilePath

On Error Resume Next

	mblnError = False
	mstrAction = Request.Form("Action")
	mstrFilePath = Request.Form("FilePath")


	'Establish the absolute path to the database
	Set conn = Server.CreateObject("ADODB.Connection")
	If len(mstrFilePath) > 0 then
		strDBpath = mstrFilePath
	ElseIf len(session("DSN_NAME")) > 0 then
		conn.Open session("DSN_NAME")
		If conn.State = 1 then
			strDBpath = conn.DefaultDatabase & ".mdb"
			conn.Close
		End If
	Else
		mblnValidConnection = False
	End If

	'Test for connection to the database
	If len(strDBpath) > 0 then
		conn.Open dbProvider & "Data Source=" & strDBpath & ";"
		mblnValidConnection = (conn.State = 1)
	Else
		mblnValidConnection = False
	End If

	if mblnValidConnection and (mstrAction = "Install Upgrade") then

		'------------------------------------------------------------------------------'
		' ADD SSUsers TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "SSUsers"
		SQL = "CREATE TABLE " & strTableName & " " _
				& "(Username char (20) PRIMARY KEY," _
				& " Userpass char (20))"
		conn.Execute (SQL)
		if Err.number = 0 then
			strMessage = strMessage & "<H3><B>Table " & strTableName & " successfully added.</B></H3><BR>"
		else
			mblnError = True
			strMessage = strMessage & "<H3><Font Color='Red'>Error adding " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
		end if

		'------------------------------------------------------------------------------'
		' ADD Default Username/Password to SSUsers														   '
		'------------------------------------------------------------------------------'

		if Err.number = 0 then
			SQL = "Insert Into " & strTableName & "(Username, Userpass) Values" _
				& " ('admin','pass')"
			conn.Execute (SQL)
			if Err.number = 0 then
				strMessage = strMessage & "<H4><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Default Username: admin.</B></H4><BR><H4><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Default Password: pass.</B></H4><BR>"
			else
				mblnError = True
				strMessage = strMessage & "<H3><Font Color='Red'>Error adding default username/password</FONT></H3><BR>"	
			end if
		
		end if
		
		If not mblnError then
			strMessage = "<H3>The installation was successful.</H3><BR>" & strMessage
		Else
			strMessage = "<H3>The following error(s) occurred during the installation.</H3><BR>" & strMessage
		End If
		
	elseif mblnValidConnection and (mstrAction = "Uninstall Upgrade") then

		'------------------------------------------------------------------------------'
		' REMOVE SSUsers TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "SSUsers"
		SQL = "DROP TABLE " & strTableName
		conn.Execute (SQL)
		if Err.number = 0 then
			strMessage = strMessage & "<H3><B>Table " & strTableName & " successfully removed.</B></H3><BR>"
		else
			mblnError = True
			strMessage = strMessage & "<H3><Font Color='Red'>Error removing " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
		end if

		If not mblnError then
			strMessage = "<H3>The uninstall was successful.</H3><BR>" & strMessage
		Else
			strMessage = "<H3>The following error(s) occurred during the uninstall.</H3><BR>" & strMessage
		End If
	End If

	On Error Resume Next
	
	conn.Close 
	Set conn = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<TITLE>Sandshot Sofware Database Upgrade Utility</TITLE>
</HEAD>
<BODY>
<P><H2>Sandshot Sofware Database Upgrade Utility</H2>
<P></P>
<P><H4>Welcome to the WebStore Manager Add-on upgrade utility for Lagarde's Storefront 5.0, Microsoft 
Access or SQL Server database.</H4>
<P>This utility makes the changes required to use the integrated security with Sandshot Software's WebStore 
Manager add-on for Lagarde's Storefront 5.0 database (all versions.)&nbsp; It accomplishes the following actions.</P>
<OL>
  <LI>creates a new table SSUsers and creates a default username/password
</OL>
<P>Instructions for use:</P>
<OL>
  <LI>This file must be located in your active Storefront 
  web. 
  <LI>Run it from your web browser.</LI></OL>
  
<P>Disclaimer: This utility is provided without warranty. While it has been 
successfully tested using the standard Storefront database (Access 97 and Access 
2000 versions), no guarantee regarding fitness for use in your application is 
made. Always make a backup of your database prior to making changes to it.</P>
<form action="ssWebStoreMgrSF5_DBUpgradeTool.asp" method="POST">
<% 
If len(session("DSN_NAME"))=0 or not mblnValidConnection then %>
<b>Select database to upgrade: </b><input type="file" name="FilePath">
<% 
	If len(mstrFilePath)=0 and not len(mstrAction)=0 then
		Response.Write "<Font Color='Red'>&nbsp;Please select a database.</Font>" 
	elseif not mblnValidConnection and len(mstrAction)<>0 then 
		Response.Write "<Font Color='Red'>" & strDBpath & " is not a valid Microsoft Access database.</FONT>" 
	End If
Else
	Response.Write "<h4>Database to be upgraded: <i>" & strDBpath & "</i></h4>"
End If %>

<p><input type=submit value="Install Upgrade" id=submit1 name=Action></p>
<p><input type=submit value="Uninstall Upgrade" id=submit2 name=Action></p>
</form>

<P><%= strMessage %></P>

<p><a href="/admin.asp">Return to Main Admin Page</a></p>

</BODY></HTML>
