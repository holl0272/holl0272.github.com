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
<form name="frmPassword" id="frmPassword" onsubmit="return ValidInput(this);" method="post" action="<%= mstrLoginPageName %>" style="display:inline">
	<input type="hidden" name="Action" ID="Action" value="Login">
	<input type=hidden name="PrevPage" ID="PrevPage" value="<%= mstrPrevPage %>">
	<table class="Section" border="0" cellpadding="0" cellspacing="0" width="95%">
		<tr>
			<td width="100%" align="center">
				<p>&nbsp;</p>
				<table class="Section" border="1" cellspacing="0" cellpadding="0" width="100%">
					<tr>
						<td class="tdMiddleTopBanner" valign="top">Customer Accounts</td>
					</tr>
					<tr>
						<td class="tdContent2" valign="top" align="left"><p>
							Returning Customers, please enter your Email Address and Password.<br />
							<font color=red>If you don't already have an account, you will have the ability to set one up at checkout.</FONT></p>
							<table border="0" cellpadding="2" style="border-collapse: collapse" width="100%">
								<tr>
									<td colspan="2" align="left"><%= strMessage %></td>
								</tr>
								<tr>
									<td align="right">Email&nbsp;Address:&nbsp;</td>
									<td align="left"><input SIZE="30" NAME="Email" ID="Email" value="<%= mstrEmail %>"></td>
								</tr>
								<tr>
									<td align="right">Password:&nbsp;</td>
									<td align="left"><input TYPE="password" SIZE="10" NAME="Password" ID="Password1"></td>
								</tr>
								<tr>
									<td align="right">&nbsp;</td>
									<td align="left"><input TYPE="checkbox" NAME="rememberMe" ID="rememberMe" VALUE="1" <% If Len(mstrEmail) > 0 Then Response.Write "checked" %>>&nbsp;<label for="rememberMe">Remember my email address</label></td>
								</tr>
								<tr>
									<td align="right">&nbsp;</td>
									<td align="left"><input type="image" class="inputImage" src="images/buttons/submit.gif" NAME="imgSubmit"></td>
								</tr>
								<tr>
									<td align="right">&nbsp;</td>
									<td align="left">
										<a href="<%= mstrLoginPageName %>?Action=EmailPwd" onclick="return forgotPassword();">Forgot your Password?</a><br />
										<a href="<%= mstrLoginPageName %>?Action=ChangePwd">Want to change your Password?</a><br />
										<!--<a href="<%= mstrLoginPageName %>?Action=createAccount">Create Account</a>-->
										<p>If you do not already have an account one will be created for you at checkout. If you wish to subscribe to our mailing list please click <a href="mailSubscribe.asp">here</a>.</p>
									</td>
								</tr>
							</table>
							&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
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

  if (isEmpty(theForm.Email,"Please enter a Email.")) {return(false);}
  if (isEmpty(theForm.Password,"Please enter a password.")) {return(false);}
  
  return(true);
}

function forgotPassword()
{
var theForm = document.frmPassword;

if (isEmpty(theForm.Email,"Please enter an email address.")){return false;}

theForm.Action.value = "EmailPwd";
theForm.submit();

return false;
}

</script>
