<%
'********************************************************************************
'*   Quick Order							                                    *
'*   Release Version:   1.0.0													*
'*   Release Date:		March 12, 2003											*
'*   Revision Date:		March 12, 2003											*
'*                                                                              *
'*   Notes: Requires Sandshot Software's Multiple Product Ordering to function  *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

%>
<script language="javascript" type="text/javascript">
<!--

	function isInteger(theField, emptyOK, theMessage)
	{
	  if (theField.value == "")
	  {
		if (emptyOK)
		{
			return(true);
		}
		{
			alert(theMessage);
			theField.focus();
			theField.select();
		  return (false);
		}
	  }

	    var i;
	    var s = theField.value;
	    for (i = 0; i < s.length; i++)
	    {
	        var c = s.charAt(i);
	        if (!((c >= "0") && (c <= "9")))
	        {
				alert(theMessage);
				theField.focus();
				theField.select();
	            return (false);
	        }
	    }
	  if (s > 0)
	  {
	  return (true);
	  }
	  {
		alert(theMessage);
		theField.focus();
		theField.select();
		return (false);
	  }
	}
// -->
</script>
<!--
Quick Order StoreFront 5.0 add-on by Sandshot Software
Requires <a href="http://www.sandshot.net/">Sandshot Software</a> Multiple Product Ordering to function
-->
			<form action="addproduct.asp" method="POST" ID="Form1">
			<input type="hidden" name="QuickOrder" ID="QuickOrder" value="1">
      		<table border="0" cellpadding="3" cellspacing="0">
					<tr>
						<td width="40"></td>
						<td align = center>
							<font face="Verdana,Arial" color="#000080" size="2">
								<b>Product ID</b>
							</font></td>
						<td align = center>
							<font face="Verdana,Arial" color="#000080" size="2">
								<b>Quantity</b>
							</font>
						</td>
						<td width="40"></td>
					</tr>
<%
Dim mlngCounter
If cblnDisplayProductIDs Then
	Dim mobjRSProducts
	Dim mstrSelect
	Set mobjRSProducts = CreateObject("ADODB.RECORDSET")
	mobjRSProducts.CursorLocation =	3			'adUseClient
	mobjRSProducts.Open "Select prodID, prodName from sfProducts where prodEnabledIsActive=1 Order By prodID",cnn,3,3
	If mobjRSProducts.EOF Then
		mstrSelect = "<input type=text name=PRODUCT size=40>"
	Else
		mstrSelect = "<select name=PRODUCT>"
		mstrSelect = mstrSelect & "<option value=''>Select a product</option>"
		Do While Not mobjRSProducts.EOF
			mstrSelect = mstrSelect & "<option value='" & Trim(mobjRSProducts.Fields("prodID").Value) & "'>" & Server.HTMLEncode(Trim(mobjRSProducts.Fields("prodID").Value)) & " - " & Server.HTMLEncode(Trim(mobjRSProducts.Fields("prodName").Value)) & "</option>"
			mobjRSProducts.MoveNext
		Loop
		mstrSelect = mstrSelect & "</select>"
	End If
	mobjRSProducts.Close
	Set mobjRSProducts = Nothing
End If
%>
					<% for mlngCounter = 1 to clngNumEntries %>
						<tr>
							<td></td>
							<td align = center>
								<% If cblnDisplayProductIDs Then %>
								<%= mstrSelect %>
								<% Else %>
								<input type=text name=PRODUCT ID="PRODUCT<%= mlngCounter %>" size=40>
								<% End If %>
							</td>
							<td align = center>
								<input type=text name=QUANTITY ID="QUANTITY<%= mlngCounter %>" size=4 onblur="return isInteger(this, true, 'Please enter an integer greater than one for the quantity')">
							</td>
							<td></td>
						</tr>
					<% next %>
					<tr>
						<td colspan = 4 align = center>
							<input type="image" class="inputImage" name="AddProduct" ID="AddProduct" src="<%= C_BTN03 %>" alt="Add To Cart">
						</td>
					</tr>
			</table>
				</form>
