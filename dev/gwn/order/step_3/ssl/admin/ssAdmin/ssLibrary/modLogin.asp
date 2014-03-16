<%
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version   1.1                                                      *
'*   Release Date      February 15, 2002			                            *
'*                                                                              *
'*   Release 1.1                                                                *
'*     - bug fix - corrected password changing routine                          *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Function ValidUserName()
'assumes the username/password is in the request.form object

dim pstrUserName, pstrPassword
dim sql, pbojRS

	pstrUserName = trim(Request.Form("UserName"))
	pstrPassword = trim(Request.Form("Password"))
	
	If len(pstrUserName) = 0 then
		If len(pstrPassword) = 0 then
			ValidUserName = "<h3>Please Login in.</h3>"
			Exit Function
		Else
			ValidUserName = "<h3><b>Please enter a username</b></h3>"
			Exit Function
		End If
	Else
		If len(pstrPassword) = 0 then
			ValidUserName = "<h3><b>Please enter a password</b></h3>"
			Exit Function
		End If
	End If

'On Error Resume Next

	set pbojRS = GetRS("Select Userpass from SSUsers where UserName = '" & pstrUserName & "'")
	
	If pbojRS.State <> 1 Then
		ValidUserName = "<div class='FatalError'>You need to upgrade your database to use integrated security</div>" _
					  & "<h3><a href='ssInstallationPrograms/ssWebStoreMgrSF5_DBUpgradeTool.asp'>Click here to upgrade</a></h3>"
	Else
		if pbojRS.eof or pbojRS.bof then
			ValidUserName = "<h3><b>You entered an invalid username. Please try again.</b></h3>"
		elseif trim(pbojRS("Userpass")) = pstrPassword then
			ValidUserName = "True"
			session("login") = pstrUserName
		else
			ValidUserName = "<h3><b>You entered an invalid password. Please try again.</b></h3>"
			Exit Function
		end if
	End If
	
End Function

Function ChangePassword()

dim pstrUserName, pstrUserName1, pstrPassword, pstrPassword1, pstrPassword2
dim sql, pbojRS

	pstrUserName = trim(Request.Form("UserName"))
	pstrPassword = trim(Request.Form("Password"))
	pstrUserName1 = trim(Request.Form("UserName1"))
	pstrPassword1 = trim(Request.Form("Password1"))
	pstrPassword2 = trim(Request.Form("Password2"))

	If len(pstrUserName) = 0 then
		ChangePassword = "<h3><b>Please enter a username</b></h3>"
		Exit Function
	Elseif len(pstrPassword) = 0 then
		ChangePassword = "<h3><b>Please enter your current password</b></h3>"
		Exit Function
	Elseif len(pstrUserName1) = 0 then
		ChangePassword = "<h3><b>Please enter a new username</b></h3>"
		Exit Function
	Elseif len(pstrPassword1) = 0 then
		ChangePassword = "<h3><b>Please enter the new password</b></h3>"
		Exit Function
	Elseif len(pstrPassword2) = 0 then
		ChangePassword = "<h3><b>Please re-type your password</b></h3>"
		Exit Function
	Elseif (pstrPassword2 <> pstrPassword1) then
		ChangePassword = "<h3><b>The passwords you entered do not match.</b></h3>"
		Exit Function
	End If

'On Error Resume Next

	set pbojRS = GetRS("Select Userpass from SSUsers where UserName = '" & pstrUserName & "'")
	
	If Err.number<> 0 Then
		If instr(1,Err.Description,"cannot find the input table or query 'SSUsers'") <> 0 Then
			ChangePassword = "<div class='FatalError'>You need to upgrade your database to use integrated security</div>"
		Else
			ChangePassword = "<div class='FatalError'>Error: " & Err.number & " - " & Err.Description & "</div>"
		End If
	Else
		if pbojRS.eof or pbojRS.bof then
			ChangePassword = "<h3><b>You entered an invalid username. Please try again.</b></h3>"
		elseif trim(pbojRS("Userpass")) = pstrPassword then
			sql = "Update SSUsers set Userpass = '" & pstrPassword1 & "', [Username]='" & pstrUserName1 & "' where [Username] = '" & pstrUserName & "'"
			cnn.Execute sql,,128
			session("login") = pstrUserName
			ChangePassword = "Username/Password Successfully Changed"
'			ChangePassword = "<b>This functionality has been disabled.</b> The Username/Password would have been changed"
		else
			ChangePassword = "<h3><b>You entered an invalid password. Please try again.</b></h3>"
		end if
	End If
	
	set pbojRS = nothing

End Function
%>
<% Sub ShowLoginForm(strMessage) %>
<script language='javascript'>
function isEmpty(theField,theMessage)
{
if (theField.value == "")
	{	
	alert(theMessage);
	theField.focus();
	theField.select();
	return(true);
	}
	{
	return(false);
	}
}

function ValidInput(theForm)
{

  if (isEmpty(theForm.UserName,"Please enter a username.")) {return(false);}
  if (isEmpty(theForm.Password,"Please enter a password.")) {return(false);}
  
  return(true);
}
</script>
<center>
<form id=form1 name=form1 onsubmit='return ValidInput(this);' method=post>
<input type=hidden name='Action' value='Login'>
<input type=hidden name='PrevPage' value='<%= mstrPrevPage %>'>
<table class="tbl" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="3" align=center><span id="spanDetailTitle">Login</span></th>
  </tr>
    <tr>
      <td width="100%" align=center><p><%= strMessage %></p></td>
    </tr>
    <tr>
      <td width="100%" align=center>Username:&nbsp;<input type="text" name="UserName" size="20" value="admin"></td>
    </tr>
    <tr>
      <td width="100%" align=center>Password:&nbsp;<input type="password" name="Password" size="20" value="pass"></td>
    </tr>
    <tr>
      <td width="100%" align=center></td>
    </tr>
    <tr>
      <td width="100%" align=center><input class='butn' type="submit" value="Submit" name="B1"></td>
    </tr>
  </table>
</form>
</center>
</body>
</html>
<% End Sub 'ShowLoginForm %>
<% Sub ShowChangePasswordForm(strMessage) %>
<script language='javascript'>
function isEmpty(theField,theMessage)
{
if (theField.value == "")
	{	
	alert(theMessage);
	theField.focus();
	theField.select();
	return(true);
	}
	{
	return(false);
	}
}

function ValidInput(theForm)
{

  if (isEmpty(theForm.UserName,"Please enter a username.")) {return(false);}
  if (isEmpty(theForm.Password,"Please enter a password.")) {return(false);}
  if (isEmpty(theForm.UserName1,"Please enter a username.")) {return(false);}
  if (isEmpty(theForm.Password1,"Please enter a password.")) {return(false);}
  if (isEmpty(theForm.Password2,"Please enter a password.")) {return(false);}
  
  return(true);
}
</script>
<center>
<form id=form1 name=form1 onsubmit='return ValidInput(this);' method=post>
<input type=hidden name='Action' value='ChangePwd'>
<table class="tbl" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="3" align=center><span id="spanDetailTitle">Change Password</span></th>
  </tr>
    <tr>
      <td align=center colspan=2><p><%= strMessage %></p></td>
    </tr>
    <tr>
      <td>Username:&nbsp; </td><td><input type="text" name="UserName" size="20" value='<%= session("login") %>'></td>
    </tr>
    <tr>
      <td>Current Password:&nbsp; </td><td><input type="password" name="Password" size="20"></td>
    </tr>
    <tr>
      <td>New Username:&nbsp; </td><td><input type="text" name="UserName1" size="20" value='<%= session("login") %>'></td>
    </tr>
    <tr>
      <td>New Password:&nbsp; </td><td><input type="password" name="Password1" size="20"></td>
    </tr>
    <tr>
      <td>Re-Type Password:&nbsp; </td><td><input type="password" name="Password2" size="20"></td>
    </tr>
    <tr>
      <td></td>
    </tr>
    <tr>
      <td align=center colspan=2><input class='butn' type="submit" value="Submit" name="B1"></td>
    </tr>
  </table>
</form>
<p><a href="Admin.asp">Return To Admin Menu</a></p>
</center>
</body>
</html>
<% End Sub 'ShowChangePasswordForm %>