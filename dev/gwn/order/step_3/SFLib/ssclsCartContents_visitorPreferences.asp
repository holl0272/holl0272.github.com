<%
'********************************************************************************
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************
'
'	This page is separated from ssclsCartContents.asp to enable easier editing
'   It is located inside the class module, Public Sub displayVisitorShippingPreferences
'
'	The html fragment on this page MUST be a self-contained table
'
'**********************************************************
'*	Sub routine variables
'**********************************************************
	
	Dim i
	
	%>
	    <form name="frmVisitorShippingPreferences" id="frmVisitorShippingPreferences" action="order.asp" method="post" onsubmit="return validShippingPreference(this)">
		<table id="divShippingHandlingCalculatorWrapper" class="Section" cellpadding="0" cellspacing="0" border="1">
		  <tr>
		    <td>
		      <table class="Section" cellpadding="5" cellspacing="0" border="0" align="center" width="100%">
			    <tr>
			      <th class="tdTopBanner" colspan="2" nowrap>Shipping and handling calculator:</th>
			    </tr>
			    <tr>
				  <td colspan="2">To calculate your shipping, handling, and tax costs please enter the destination you would like to ship to below.<br /></td>
				</tr>
				<tr>
				  <td><b>State:</b><br /><select name="visitorState" id="visitorState"><%= getStateList(pstrState) %></select></td>
				  <td rowspan="3"><a href="#" onClick="window.open('ssl/viewUPSMap.asp','IWIN', 'status=no,location=no,menu=no,scrollbars,width=700,height=600,');"><img border="0" src="ssl/images/upslink.gif" align="absbottom" alt="View UPS delivery times"></a></td>
				</tr>
				<tr>
				  <td><b>Zip Code:</b><br /><input type="text" name="visitorZIP" id="visitorZIP" size="10" maxlength="10" value="<%= pstrZIP %>"></td>
				</tr>
				<tr>
				  <td><b>Country</b><br /><select name="visitorCountry" id="visitorCountry" onchange="visitorCountry_onchange(this);"><%= getCountryList(pstrCountry, adminOriginCountry) %></select></td>
				</tr>
				<%
				If isArray(paryAvailableShippingMethods) Then
					If UBound(paryAvailableShippingMethods) >= 0 Then
				%>
				<tr>
					<td><hr /></td>
				</tr>
				<tr>
					<td align="center">
					  <table class="Section" border="1" cellpadding="2" cellspacing="0">
					    <tr><th colspan="2" class="tdTopBanner" nowrap>Available Shipping Options</th></tr>
					    <tr><td colspan="2" style="text-align: left; padding-left: 3pt;">Please select your preferred shipping method and press update to change your cart total.</td></tr>
						<%
						For i = 0 To UBound(paryAvailableShippingMethods)
							If CStr(paryAvailableShippingMethods(i)(0)) = CStr(pstrShipMethodCode & "") Then
								Response.Write "<tr><td class=tdContent align=left><input type=radio name=visitorPreferredShippingCode id=visitorPreferredShippingCode" & i & " value=" & paryAvailableShippingMethods(i)(0) & " checked>&nbsp;<label for=visitorPreferredShippingCode" & i & "><strong>" & paryAvailableShippingMethods(i)(1) & "</strong></label></td><td class=tdContent align=right><strong>" & customCurrency(paryAvailableShippingMethods(i)(2)) & "</strong></td></tr>"
							Else
								Response.Write "<tr><td class=tdContent align=left><input type=radio name=visitorPreferredShippingCode id=visitorPreferredShippingCode" & i & " value=" & paryAvailableShippingMethods(i)(0) & ">&nbsp;<label for=visitorPreferredShippingCode" & i & ">" & paryAvailableShippingMethods(i)(1) & "</label></td><td class=tdContent align=right>" & customCurrency(paryAvailableShippingMethods(i)(2)) & "</td></tr>"
							End If
						Next 'i
						
						%>
					  </table>
					</td>
				</tr>
				<%
					End If
				End If	'isArray(paryAvailableShippingMethods)
				%>
				<tr>
				  <td align="center"><input type="submit" name="updateVisitorShippingPreferences" ID="updateVisitorShippingPreferences" value="Update" /></td>
			    </tr>
		      </table>
		    </td>
		  </tr>
		</table>
	    </form>
