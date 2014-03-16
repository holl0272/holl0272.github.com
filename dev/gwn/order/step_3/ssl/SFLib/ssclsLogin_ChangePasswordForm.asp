<%
'********************************************************************************
'*   Page Protector for StoreFront 5.0                                          *
'*   Release Version:   1.00.002                                                *
'*   Included with ssclsLogin.asp v1.00.002										*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<script language="javascript" type="text/javascript">
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

  if (isEmpty(theForm.Login,"Please enter a Email.")) {return(false);}
  if (isEmpty(theForm.Password,"Please enter a password.")) {return(false);}
  if (isEmpty(theForm.NewPassword1,"Please enter a password.")) {return(false);}
  if (isEmpty(theForm.NewPassword2,"Please enter a password.")) {return(false);}
  if (theForm.NewPassword1.value != theForm.NewPassword2.value)
  {
	alert("Your passwords do not match.");
	theForm.NewPassword1.focus();
	theForm.NewPassword1.select();
	return(false);
  }
  
  return(true);
}
</script>
	<br />
	<form id="form1" name="form1" onsubmit="return ValidInput(this);" method="post" action="<%= mstrLoginPageName %>">
		<input type="hidden" name="Action" value="ChangePwd" ID="ChangePwd">
		<table class="tbl" cellpadding="3" cellspacing="0" border="0">
			<colgroup>
				<col align="right">
				<col align="left">
			</colgroup>
					<tr class="tblhdr">
						<th class="tdMiddleTopBanner" colspan="2" align="center">Change Password</th>
					</tr>
					<tr>
						<td align="center" colspan="2"><p><%= strMessage %></p>
						</td>
					</tr>
					<tr>
						<td align="right">Email:&nbsp;</td>
						<td><input type="text" name="Login" ID="Login" size="30" value="<%= session("login") %>"></td>
					</tr>
					<tr>
						<td align="right">Current Password:&nbsp;</td>
						<td><input type="password" name="Password" ID="Password" size="20"></td>
					</tr>
					<tr>
						<td align="right">New Password:&nbsp;</td>
						<td><input type="password" name="NewPassword1" ID="NewPassword1" size="20"></td>
					</tr>
					<tr>
						<td align="right">Re-Type Password:&nbsp;</td>
						<td><input type="password" name="NewPassword2" ID="NewPassword2" size="20"></td>
					</tr>
					<tr>
						<td></td>
					</tr>
					<tr>
						<td align="center" colspan="2"><input class="butn" type="submit" value="Submit" name="Submit1" ID="Submit1"></td>
					</tr>
		</table>
	</form>
<% If isLoggedIn Then %>
	<p><a href="myAccount.asp">Return to myAccount</a></p>
<% End If %>