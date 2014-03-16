<%Option Explicit 
Server.ScriptTimeout = 300

	sub debugprint(sField1,sField2)
	Response.Write "<H3>" & sField1 & ": " & sField2 & "</H3><BR>"
	end sub

	Sub UpdateAction(strMessage)
		Response.Write "<script>UpdateAction('" & strMessage & "');</script>" & vbcrlf
	End Sub

'***********************************************************************************************

	Function TableExists(objConn,strTableName)

	Dim prsTest

	On Error Resume Next

		Set prsTest = Server.CreateObject("ADODB.RECORDSET")
		With prsTest
			.CursorLocation = adUseClient
			.MaxRecords = 1
			.Open strTableName, objConn, adOpenStatic, adLockReadOnly, adCmdTable
			TableExists = (Err.number = 0)
			.Close
		End With
		Set prsTest = Nothing

	End Function	'TableExists

'***********************************************************************************************

	Function OpenDatabase(objCnn, strFilePath, blnDSN)

	On Error Resume Next

		Set objCnn = Server.CreateObject("ADODB.Connection")
		If blnDSN Then
			objCnn.Open strFilePath
		Else
			objCnn.Open dbProvider & "Data Source=" & strFilePath & ";"
		End If
		If Err.number <> 0 Then
			debugprint Err.number,Err.Description
		End If
		OpenDatabase = (objCnn.State = 1)

	End Function	'OpenDatabase

'***********************************************************************************************

	Sub ImportDatabase
	
	Dim pobjcnnImport
	Dim prsTarget
	Dim prsSource
		
	Dim pstrFilePath
	Dim pstrTableName
	Dim pstrFieldName_City
	Dim pstrFieldName_County
	Dim pstrFieldName_PostalCode
	Dim pstrFieldName_TaxRate
	Dim pstrFieldName_State
	Dim pblnAppend
	Dim pblnSuccess
	
		pblnSuccess = True
		
		pstrFilePath = Trim(Request.Form("ImportPath"))
		pstrTableName = Trim(Request.Form("Table"))
		pstrFieldName_City = Trim(Request.Form("City"))
		pstrFieldName_County = Trim(Request.Form("County"))
		pstrFieldName_PostalCode = Trim(Request.Form("PostalCode"))
		pstrFieldName_State = Trim(Request.Form("State"))
		pstrFieldName_TaxRate = Trim(Request.Form("TaxRate"))
		pblnAppend = (Request.Form("overwrite") = 1)

'debugprint "pstrFilePath",server.MapPath("/")
'debugprint "pstrFilePath",pstrFilePath
'debugprint "pstrTableName",pstrTableName
'debugprint "pstrFieldName_City",pstrFieldName_City
'debugprint "pstrFieldName_County",pstrFieldName_County
'debugprint "pstrFieldName_PostalCode",pstrFieldName_PostalCode
'debugprint "pstrFieldName_State",pstrFieldName_State
'debugprint "pstrFieldName_TaxRate",pstrFieldName_TaxRate
'debugprint "pblnAppend",pblnAppend

		If OpenDatabase(pobjcnnImport,pstrFilePath,False) Then
	
			If Not pblnAppend Then 
				UpdateAction "Deleting existing tax rates"
				conn.Execute "Delete from ssTaxTable",,128
			End If
			
			'Open the source table
			Set prsSource = server.CreateObject("ADODB.RECORDSET")
			prsSource.CursorLocation =	adUseClient
			prsSource.Open pstrTableName, pobjcnnImport, adOpenStatic, adLockOptimistic, adCmdTable
			If cBool(prsSource.State) Then
				UpdateAction pstrTableName & " successfully opened"
			Else
				pblnSuccess = False
				If Err.number = 0 Then
					UpdateAction "<font color=red><b>Error opening " & pstrTableName & "</b></font>"
				Else
					UpdateAction "<font color=red><b>Error opening " & pstrTableName & " " & Err.number & ":" & Err.Description & "</b></font>"
				End If
			End If
			debugprint "prsSource.RecordCount",prsSource.RecordCount
			
			'Open the target table
			Set prsTarget = server.CreateObject("ADODB.RECORDSET")
			prsTarget.CursorLocation =	adUseClient
			prsTarget.CacheSize = 100
			prsTarget.Open "ssTaxTable", conn, adUseServer, adLockOptimistic, adCmdTable
			If cBool(prsSource.State) Then
				UpdateAction pstrTableName & " successfully opened"
			Else
				pblnSuccess = False
				If Err.number = 0 Then
					UpdateAction "<font color=red><b>Error opening ssTaxTable</b></font>"
				Else
					UpdateAction "<font color=red><b>Error opening ssTaxTable" & Err.number & ":" & Err.Description & "</b></font>"
				End If
			End If
			
			debugprint "prsTarget.RecordCount",prsTarget.RecordCount
'On Error Resume Next

			If pblnSuccess Then

				UpdateAction "Inserting new tax rates"
				If mblnTurnStatusOn Then Response.Write "<script>UpdateRecordCount(' of " & prsSource.RecordCount & " records to insert');</script>"
				For i=1 to prsSource.RecordCount
					prsTarget.AddNew

					If len(pstrFieldName_City) > 0 Then prsTarget.Fields("City").Value = prsSource.Fields(pstrFieldName_City).Value
					If len(pstrFieldName_County) > 0 Then prsTarget.Fields("County").Value = prsSource.Fields(pstrFieldName_County).Value
					If len(pstrFieldName_PostalCode) > 0 Then prsTarget.Fields("PostalCode").Value = prsSource.Fields(pstrFieldName_PostalCode).Value
					If len(pstrFieldName_State) > 0 Then prsTarget.Fields("LocaleAbbr").Value = prsSource.Fields(pstrFieldName_State).Value
					If len(pstrFieldName_TaxRate) > 0 Then 
						prsTarget.Fields("TaxRate").Value = prsSource.Fields(pstrFieldName_TaxRate).Value
					Else
						prsTarget.Fields("TaxRate").Value = 0
					End If
					
					If Err.number <> 0 Then
						pblnSuccess = False
						UpdateAction "<font color=red><b>Inserting new tax rates failed.<br>Error " & Err.number & ":" & Err.Description & "</b></font>"
'						Response.Write "<script>UpdateRecordCount('');</script>"
'						Response.Write "<script>UpdatePosition('<font color=red><b>Error " & Err.number & ":" & Err.Description & "</b></font>');</script>"
						Exit For
					End If				
					If mblnTurnStatusOn Then Response.Write "<script>UpdatePosition('" & i & "');</script>"
					prsSource.MoveNext
				Next
On Error Goto 0						
				If pblnSuccess Then prsTarget.UpdateBatch
			
On Error Resume Next
			End If
			
			prsTarget.Close
			Set prsTarget = Nothing

			prsSource.Close
			Set prsSource = Nothing
			
			pobjcnnImport.Close
			Set pobjcnnImport = Nothing
			
			If pblnSuccess Then UpdateAction "Database imported"

		End If
		
	End Sub	'ImportDatabase

'***********************************************************************************************

Const dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Const mblnTurnStatusOn = True

dim conn,sql,strTableName
Dim mblnSQLServer
Dim mblnValidConnection
Dim mblnTableExists
Dim mstrDBPath
Dim mstrFilePath
Dim mstrMessage
dim mstrAction
dim mblnError
dim strSQL
dim i

On Error Resume Next

	mblnError = False
	mstrAction = Request.Form("Action")
	mstrFilePath = Request.Form("FilePath")
	mblnSQLServer = (LCase(Request.Form("SQLServer")) = "on")
	
	If len(mstrFilePath) > 0 then
		mstrDBPath = mstrFilePath
		mblnValidConnection = OpenDatabase(conn,mstrFilePath,False)
	ElseIf len(Application("DSN_NAME")) > 0 then
		mblnValidConnection = OpenDatabase(conn,Application("DSN_NAME"),True)
		If mblnValidConnection then mstrDBPath = conn.DefaultDatabase & ".mdb"
	Else
		mblnValidConnection = False
	End If

Select Case mstrAction
	Case "Import"
		If mblnValidConnection Then	mblnTableExists = TableExists(conn,"ssTaxTable")
		Call WritePage
		Call ImportDatabase
	Case "Install"

		'------------------------------------------------------------------------------'
		' ADD ssShipZones TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssTaxTable"
		If mblnSQLServer Then
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(TaxRateID Int Identity," _
					& " City char (50),"_
					& " County char (50),"_
					& " PostalCode char (10),"_
					& " LocaleAbbr char (3),"_
					& " TaxRate Decimal Not Null)"
		Else
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(TaxRateID Counter PRIMARY KEY," _
					& " City char (50),"_
					& " County char (50),"_
					& " LocaleAbbr char (3),"_
					& " PostalCode char (10),"_
					& " TaxRate Single Not Null)"
		End If
		conn.Execute (SQL)

		mblnTableExists = (Err.number = 0)
		if mblnTableExists then
			mstrMessage = mstrMessage & "<H3><B>" & strTableName & " table successfully added.</B></H3><BR>"
		else
			mblnError = True
debugprint "sql",sql
			mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error adding " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
			mblnTableExists = (Instr(1,Err.Description,"'ssTaxTable' already exists")>0)
			Err.Clear
		end if
		Call WritePage	
	Case "Uninstall"

		'------------------------------------------------------------------------------'
		' REMOVE ssShipZones TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssTaxTable"
		SQL = "DROP TABLE " & strTableName
		conn.Execute (SQL)

		mblnTableExists = NOT (Err.number = 0)
		if NOT mblnTableExists then
			mstrMessage = mstrMessage & "<H3><B>" & strTableName & " table successfully removed.</B></H3><BR>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error removing " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
			mblnTableExists = (Instr(1,Err.Description,"'ssTaxTable' does not exist")=0)
			Err.Clear
		end if
		Call WritePage	
	Case Else
		If mblnValidConnection Then	mblnTableExists = TableExists(conn,"ssTaxTable")
		Call WritePage	
End Select

conn.Close 
Set conn = Nothing

Sub WritePage
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<TITLE>Sandshot Software - Tax Manager Database Upgrade Utility</TITLE>
<script language="javascript">

function UpdateAction(strAction){ document.all.divAction.innerHTML = strAction;}

function UpdatePosition(strAction){ document.all.divPosition.innerHTML = strAction;}

function UpdateRecordCount(strAction){ document.all.divRecordCount.innerHTML = strAction;}

</script>
</HEAD>
<BODY>
<P><H2>Sandshot Sofware Database Upgrade Utility</H2>
<P></P>
<P><H4>Welcome to the Tax Manager Module upgrade utility for Lagarde's Storefront database.</H4>
<P>This utility makes the changes required to use Sandshot Software's Tax Manager 
add-on for Lagarde's Storefront. It accomplishes the following actions.</P>

<h4>Creates the ssTaxRates table</h4>

<P>Instructions for use:</P>
<OL>
  <LI>This file must be located in your active Storefront web. 
  <LI>Run it from your web browser.</LI></OL>
  
<P>Disclaimer: This utility is provided without warranty. While it has been 
successfully tested using the standard Storefront databases (Access 97, Access 
2000, and SQL Server versions), no guarantee regarding fitness for use in your appication is 
made. Always make a backup of your database prior to making changes to it.</P>

<% If len(mstrMessage) > 0 Then Response.Write "<p>" & mstrMessage & "</p>" %>

<form action="" method="POST" id=form1 name=form1>

<% If mblnTableExists Then %>
<table border=1 cellpadding=2 cellspacing=0 width=95% ID="Table1">
<tr>
<td colspan=3><h4>Import Tax Information from a database</h4>
<i>To import tax rate information from an external database fill in the information below. The form
is prefilled with the default database table structure. The imported database can be of different 
structure, but you need to fill in the appropriate table and field names. If the imported database 
does not contain all of the fields, just leave the corresponding field name blank.</i> 
</td>
</tr>
<tr>
<td colspan=3><div id="divAction">&nbsp;</div></td>
</tr>
<tr>
<td colspan=3><div id="divPosition">&nbsp;</div><div id="divRecordCount"></div>&nbsp;</td>
</tr>

<tr>
<td colspan=3><i>Database structure</i></td>
</tr>
<tr>
<td><b>&nbsp;</b></td><td>Imported Database</td><td>StoreFront Database</td>
</tr>
<tr>
<td align=right>Database:</td><td><input type="file" name="ImportPath" ID="ImportPath"></td><td><i><%= mstrDBPath %></i></td>
</tr>
<tr>
<td align=right>Table:</td><td><input name="Table" value="ssTaxTable" ID="Text1"></td><td><i>ssTaxTable</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="City" value="City" ID="Text2"></td><td><i>City</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="County" value="County" ID="Text3"></td><td><i>County</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="PostalCode" value="PostalCode" ID="Text4"></td><td><i>PostalCode</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="TaxRate" value="TaxRate" ID="Text5"></td><td><i>TaxRate</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="State" value="LocaleAbbr" ID="Text6"></td><td><i>State</i></td>
</tr>
<tr>
<td>&nbsp;</td>
<td colspan=2>
	<input type=radio name="overwrite" value=0 checked ID="Radio1">&nbsp;Overwrite existing data.<br>
	<input type=radio name="overwrite" value=1 ID="Radio2">&nbsp;Append to existing data.
</td>
</tr>
<tr>
<td>&nbsp;</td>
<td colspan=2><input type=submit value="Import" id="Action_Import" name="Action"></td>
</tr>
</table>
<p><hr></p>
<table border=1 cellpadding=2 cellspacing=0 width=95% ID="Table2">
<tr>
<td>Remove the TaxMgrUpgrade from the database&nbsp;&nbsp;<input type=submit value="Uninstall" id="Submit1" name="Action"></td>
</tr>
</table>

<% Else %>
<%   If not mblnValidConnection then %>
<b>Select database to upgrade: </b><input type="file" name="FilePath" ID="File2">
<% 
	If len(mstrFilePath)=0 and not len(mstrAction)=0 then
		Response.Write "<Font Color='Red'>&nbsp;Please select a database.</Font>" 
	elseif not mblnValidConnection and len(mstrAction)<>0 then 
		Response.Write "<Font Color='Red'>" & mstrDBPath & " is not a valid Microsoft Access database.</FONT>" 
	End If
Else
	Response.Write "<h4>Database to be upgraded: <i>" & mstrDBPath & "</i></h4>"
End If %>
<p><input type=checkbox name="SQLSERVER" ID="SQLSERVER">&nbsp;This is a SQL Server database</p>
<p><input type=submit value="Install" id="Action_Install" name="Action"></p>
<% End If %>
</form>

</BODY></HTML>
<%
End Sub 'WritePage
%>