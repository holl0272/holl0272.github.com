<%
'********************************************************************************
'*   Promotion Manager for StoreFront 5.0										*
'*   Release Version:	2.00.002 												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		September 5, 2003										*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************
'Const cstrFontColor = "white"
Const cstrFontColor = ""

'************************************************************************************************

Sub DisplayError(strMessage)
%>
<script language="javascript" type="text/javascript">
var blnPromotionSubmitted = false;
function SubmitPromotion(theForm)
{
var theButton = theForm.btnSubmit;
	
	if (blnPromotionSubmitted){return false;}
	
	if (theForm.PromoCode.value == "")
	{
		alert("Please enter a discount code.")
		theForm.PromoCode.focus();
		return false;
	}
	
	theButton.value = "Retrieving code . . .";
	
	blnPromotionSubmitted = true;
	theForm.submit();
	return false;
	
}
	
</script>
<form ID="frmPromotionRegister" Name="frmPromotionRegister" action="PromotionRegistration.asp" method=post onsubmit="return(SubmitPromotion(this));">
<table border=0 cellPadding=1 cellSpacing=1 width=100%>
  <tr>
    <td align=center><%= mstrPromotionRegistrationMessage %></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>  
  <tr>
    <td>Discount Code:&nbsp;<input id=PromoCode name=PromoCode Value='<%= mstrPromotionCode %>'>
        &nbsp;&nbsp;&nbsp;<input id="btnSubmit" name="btnSubmit" type=submit value="Enter Discount Code">
    </td>
  </tr>
  <tr>
    <td><hr></td>
  </tr>
  <tr>
    <td align=center><a href="javascript:window.close();">Close</a></td>
  </tr>
</table>
</form>
<%
End Sub	'Display Error

'************************************************************************************************

Sub DisplaySuccess
%>
<table border=0 cellPadding=1 cellSpacing=1 width=100%>
  <tr>
    <td><%= mstrPromotionRegistrationMessage %>&nbsp;</td>
  </tr>
  <tr>
    <td><hr></td>
  </tr>
  <tr>
    <td align=center><a href="javascript:window.close();">Close</a></td>
  </tr>
</table>
<%
End Sub	'DisplaySuccess

'************************************************************************************************
%>

<table border="0" cellpadding="0" cellspacing="0" width=100% border=0>
  <tr>
    <td align=center valign=middle width=100%>
<%
'************************************************************************************************

dim aPromoItem

	If len(mstrPromotionCode) = 0 Then
		mstrPromotionRegistrationMessage = "<h3>Enter your discount code.</h3>"
		mstrPromotionRegistrationMessage = ""
		Call DisplayError(mstrPromotionRegistrationMessage)
	Else
		If mblnSuccessfulRegistration Then
			'Call DisplaySuccess
			If mblnRedirectToProduct Then
				If len(mstrPromotedProducts) > 0 Then
					aPromoItem = split(mstrPromotedProducts,";")
					mstrPromotedProducts = aPromoItem(1)
					If len(mstrPromotedProducts) > 0 Then
						Call CleanUp
						Response.Clear
						Response.Redirect "detail.asp?Product_ID=" & mstrPromotedProducts
					End If			
				End If
			End If
			
			If Len(cstrRegistrationRedirectPage) > 0 Then
				Call CleanUp
				Response.Clear
				Response.Redirect cstrRegistrationRedirectPage
			End If

			Call DisplaySuccess
		Else		'display Error Message
			Call DisplayError(mstrPromotionRegistrationMessage)
		End If
	End If
	
	Call CleanUp

'************************************************************************************************
Sub CleanUp
	On Error Resume Next
	Call cleanup_dbconnopen	'This line needs to be included to close database connection
End Sub
%>
	</td>
  </tr>
</table>

