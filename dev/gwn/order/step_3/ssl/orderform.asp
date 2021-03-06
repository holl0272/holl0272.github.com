<%
	
'@BEGINVERSIONINFO

'@APPVERSION: 50.4013.0.2

'@FILENAME: orderform.asp
	 


'@DESCRIPTION: Include File for process_order.asp

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
	
Dim sShipMethods, iStepCounter, iCustID, sCtryList, sStateList, sFormDesign, sCCList, rsProcAdmin, sMailIsActive
   	 
   	 sShipMethods = getShippingList(blnFree)
	iCustID		= Request.Cookies("sfCustomer")("custID")
	sFormDesign	= C_FORMDESIGN
	sCtryList		= getCountryList() 'C_CTRYLIST	
	Dim shipCtryList
	shipCtryList = getShipCountryList()	
	sStateList		= getStateList() 'C_STATELIST
	
	Set rsProcAdmin = Server.CreateObject("ADODB.Recordset")	
	rsProcAdmin.Open "SELECT adminSubscribeMailIsActive FROM sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
   sMailIsActive = trim(rsProcAdmin("adminSubscribeMailIsActive"))
	closeObj(rsProcAdmin)

	If bLoggedIn AND iCustID <> "" Then
		Dim rsGetCustDetails, rsGetCustShipDetails, sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2
		Dim sCustCity, sCustState, sCustStateName, sCustZip, sCustCountry, sCustCountryName, sCustPhone, sCustFax, sCustEmail, sCustCardType, sCustCardTypeName, sCustSubscribed, sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName
		Dim sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustStateName, sShipCustZip, sShipCustCountry,sShipCustCountryName, sShipCustPhone
		Dim sShipCustFax, sShipCustCompany, sShipCustEmail,sSubmitAction
	   
		' Get RecordSet of customer details and shipping details and credit card details
		Set rsGetCustDetails		= getCustomerRow(iCustID)
		Set rsGetCustShipDetails	= getCustomerShippingRow(iCustID)
		'Set rsGetCustCardDetails	= getCustomerCardRow(iCustID)
				
		' Collect billing address
		sCustFirstName		= rsGetCustDetails.Fields("custFirstName")
		sCustMiddleInitial	= rsGetCustDetails.Fields("custMiddleInitial")
		sCustLastName		= rsGetCustDetails.Fields("custLastName")
		sCustCompany		= rsGetCustDetails.Fields("custCompany")
		sCustAddress1		= rsGetCustDetails.Fields("custAddr1")
		sCustAddress2		= rsGetCustDetails.Fields("custAddr2")	   
		sCustCity			= rsGetCustDetails.Fields("custCity")
		sCustState			= rsGetCustDetails.Fields("custState")		
		sCustStateName		= getNameWithID("sfLocalesState",sCustState,"loclstAbbreviation","loclstName",1)		
		sCustZip			= rsGetCustDetails.Fields("custZip")
		sCustCountry		= rsGetCustDetails.Fields("custCountry")
		sCustCountryName	= getNameWithID("sfLocalesCountry",sCustCountry,"loclctryAbbreviation","loclctryName",1)	
		sCustPhone			= rsGetCustDetails.Fields("custPhone")
		sCustFax			= rsGetCustDetails.Fields("custFax")
		sCustEmail			= rsGetCustDetails.Fields("custEmail")
		sCustSubscribed		= rsGetCustDetails.Fields("custIsSubscribed")
	   
	    ' Change display for saved cart customers
	    If instr(1,sCustFirstName,"Saved Cart Customer",1) Then
			sCustFirstName = ""
		End If	 
	   
		' Get Ship Address
		If Not rsGetCustShipDetails.EOF Then
			sShipCustFirstName		= rsGetCustShipDetails.Fields("cshpaddrShipFirstName")
			sShipCustMiddleInitial	= rsGetCustShipDetails.Fields("cshpaddrShipMiddleInitial")
			sShipCustLastName		= rsGetCustShipDetails.Fields("cshpaddrShipLastName")
			sShipCustCompany		= rsGetCustShipDetails.Fields("cshpaddrShipCompany")
			sShipCustAddress1		= rsGetCustShipDetails.Fields("cshpaddrShipAddr1")
			sShipCustAddress2		= rsGetCustShipDetails.Fields("cshpaddrShipAddr2")	   
			sShipCustCity			= rsGetCustShipDetails.Fields("cshpaddrShipCity")
			sShipCustState			= rsGetCustShipDetails.Fields("cshpaddrShipState")
			sShipCustStateName		= getNameWithID("sfLocalesState",sShipCustState,"loclstAbbreviation","loclstName",1)	
			sShipCustZip			= rsGetCustShipDetails.Fields("cshpaddrShipZip")
			sShipCustCountry		= rsGetCustShipDetails.Fields("cshpaddrShipCountry")
			sShipCustCountryName	= getNameWithID("sfLocalesCountry",sShipCustCountry,"loclctryAbbreviation","loclctryName",1)
			sShipCustPhone			= rsGetCustShipDetails.Fields("cshpaddrShipPhone")
			sShipCustFax			= rsGetCustShipDetails.Fields("cshpaddrShipFax")
			sShipCustEmail			= rsGetCustShipDetails.Fields("cshpaddrShipEmail")
		Else
			sShipCustFirstName		= sCustFirstName
			sShipCustMiddleInitial	= sCustMiddleInitial
			sShipCustLastName		= sCustLastName
			sShipCustCompany		= sCustCompany
			sShipCustAddress1		= sCustAddress1
			sShipCustAddress2		= sCustAddress2
			sShipCustCity			= sCustCity
			sShipCustState			= sCustState
			sShipCustStateName		= sCustStateName
			sShipCustZip			= sCustZip
			sShipCustCountry		= sCustCountry
			sShipCustCountryName	= sCustCountryName
			sShipCustPhone			= sCustPhone
			sShipCustFax			= sCustFax
			sShipCustEmail			= sCustEmail
		End If 
		
	' End logged in if	
	End If
	
	' Cleanup
	closeobj(rsGetCustDetails)
	closeobj(rsGetCustShipDetails)


	' Used for iterating steps -- useful if some step is skipped
	iStepCounter = 0
	Function getStepCounter(iStepCounter)
		iStepCounter = iStepCounter + 1
		getStepCounter = iStepCounter
	End Function
	
	sSubmitAction = ""
	If NOT (bLoggedIn) OR iCustID = "" Then sSubmitAction = "this.Password.password=true;this.Password.optional = true;this.Password2.optional = true;"
	sSubmitAction = sSubmitAction & "this.Company.optional = true;this.Address2.optional = true;this.Fax.optional = true;this.Address2.optional = true;this.Instructions.optional = true;this.Email.eMail = true;this.Phone.phoneNumber = true;this.ShipState.optional = true;this.ShipCountry.optional = true;this.MiddleInitial.optional = true;"		
	If sPaymentMethod = "Credit Card" Then
		sSubmitAction = sSubmitAction & "this.CardNumber.creditCardNumber = true;this.CardExpiryMonth.creditCardExpMonth = true;this.CardExpiryYear.creditCardExpYear = true;return validate_Me(this);"
	Elseif sPaymentMethod = "PhoneFax" Then 
		sSubmitAction = sSubmitAction & "this.CardType.optional = true;this.CardName.special = true;this.CardNumber.special = true;this.CardExpiryMonth.special = true;this.CardExpiryYear.special = true;this.CheckNumber.special = true;this.BankName.special = true;this.RoutingNumber.special = true;this.CheckingAccountNumber.special = true;this.POName.special = true;this.PONumber.special = true;return validate_Me(this);"
	Else
		sSubmitAction = sSubmitAction & "return validate_Me(this);"
	End If
%>
	  <form method="post" action="verify.asp" name="form1" onSubmit="<%= sSubmitAction %>">
    <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
        <tr>
				  <td width="100%" colspan="2">
            <table border="0" width="100%" cellspacing="0" cellpadding="2">
              <tr>
                <td width="100%" class="tdContentBar"> Billing Information<span lang="en-us">&nbsp;&nbsp;
                <font size="1">(Make sure this is the address <b>CURRENTLY ON 
                FILE</b> with the Credit Card company or your order will be <b>
                DELAYED</b>)</font></span></td>
              </tr>
            </table>
          </td>
        </tr>
    
        <tr><td width="100%" colspan="2">
            <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
              <tr>
                <td nowrap align="right"><b><font color="#FF0000">*</font>First Name:</b></td>
                <td nowrap><input type="text" maxlength="50" name="FirstName" title="First Name" size="20" Style="<%= sFormDesign%>" value="<%= sCustFirstName %>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <b>MI:&nbsp;</b><input type="text" name="MiddleInitial" size="1" Style="<%= sFormDesign%>" value="<%= sCustMiddleInitial %>" maxlength="1"></td>
                <td nowrap align="right"><b><font color="#FF0000">*</font>Last
        Name:</b></td>
                <td nowrap><input type="text" maxlength="50" name="LastName" title="Last Name" size="20" Style="<%= sFormDesign%>" value="<%= sCustLastName %>"></td>
              </tr>
              <tr>
                <td nowrap align="right"><b>Company:</b></td>
                <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="Company" title="Company" size="25" value="<%= sCustCompany %>"></td>
                <td nowrap align="right">&nbsp;</td>
                <td nowrap>&nbsp;</td>
              </tr>
              <tr>
                <td nowrap align="right" height="15"></td>
                <td nowrap height="15"></td>
                <td nowrap align="right" height="15"></td>
                <td nowrap height="15"></td>
              </tr>
              <tr>
                <td nowrap align="right"><b><font color="#FF0000">*</font>Street Address:</b></td>
                <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="Address1" title="Street Address" size="25" value="<%= sCustAddress1 %>"></td>
                <td nowrap align="right"><b>Street Address 2:</b></td>
                <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="Address2" title="Address2" size="25" value="<%= sCustAddress2 %>"></td>
              </tr>
              <tr>
                <td nowrap align="right"><b><font color="#FF0000">*</font>City:</b></td>
                <td nowrap><input type="text" maxlength="50" name="City" title="City" size="20" Style="<%= sFormDesign%>" value="<%= sCustCity %>"></td>
                <td nowrap align="right"><b><font color="#FF0000">*</font>State/Province:</b></td>
                <td nowrap><select name="State" title="State" style="<%= sFormDesign%>"><option value="<%= sCustState %>"><%= sCustStateName %></option><%= sStateList %></select></td>
              </tr>
              <tr>
                <td nowrap align="right"><b><font color="#FF0000">*</font>Zip/Postal
        Code:</b></td>
                <td nowrap><input type="text" maxlength="12" name="Zip" title="Zip Code" size="6" Style="<%= sFormDesign%>" value="<%= sCustZip %>"></td>
                <td nowrap align="right"><b><font color="#FF0000">*</font>Country:</b></td>
                <td nowrap><select name="Country" title="Country" style="<%= sFormDesign%>"><option value="<%= sCustCountry %>"><%= sCustCountryName %></option><%= sCtryList %></select></td>
              </tr>
              <tr>
                <td nowrap align="right" height="15"></td>
                <td nowrap height="15"></td>
                <td nowrap align="right" height="15"></td>
                <td nowrap height="15"></td>
              </tr>
              <tr>
                <td nowrap align="right"><b> <font color="#FF0000">*</font>Phone:</b></td>
                <td nowrap><input type="text" name="Phone" maxlength="20" title="Phone Number" size="20" Style="<%= sFormDesign%>" value="<%= sCustPhone %>"></td>
                <td nowrap align="right"><b><font color="#FF0000">*</font>E-Mail:</b></td>
                <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="100" name="Email" title="Email Address" size="25" value="<%= sCustEmail %>"></td>
              </tr>
              <tr>
                <td nowrap align="right"><b>Fax:</b></td>
                <td nowrap><input type="text" name="Fax" maxlength="20" title="Fax Number" size="20" Style="<%= sFormDesign%>" value="<%= sCustFax %>"></td>
                <td nowrap align="right">&nbsp;</td>
                <td nowrap><% If sMailIsActive = "1" Then %><input type="checkbox" name="Subscribe" value="1" <%if trim(sCustSubscribed) = "1" or trim(sCustSubscribed)="" then Response.write "checked" %>><b>Add to mailing list</b><% End If %></td>
              </tr>
              <tr>
                <td width="100%" valign="top" height="20" align="left" colspan="4"></td>
              </tr>
    
            </table>
          </td></tr>
    
        <!-- Shipping Info -->     
        <tr><td width="100%" class="tdContent2">  
		    <table class="tdContent2" cellpadding="2" cellspacing="0" width="100%">
			  <tr>
			    <td width="100%" colspan="4" cellpadding="0" cellspacing="0">
			      <table border="0" cellpadding="2" width="100%" class="tdContentBar" cellspacing="0">
			        <tr>
			          <td width="100%" class="tdContentBar"> Shipping Information 
                      <font size="1">(If Different from Billing Information)</font></td>
			        </tr>
			      </table>
			    </td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>First Name:</b></td>
			    <td nowrap><input type="text" maxlength="50" name="ShipFirstName" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustFirstName %>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <b>MI:&nbsp;</b><input type="text" name="ShipMiddleInitial" size="1" Style="<%= sFormDesign%>" value="<%= sShipCustMiddleInitial %>" maxlength="1"></td>
			    <td nowrap align="right"><b>Last Name:</b></td>   
			    <td nowrap>&nbsp;&nbsp;<input type="text" maxlength="50" name="ShipLastName" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustLastName %>"></td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>Company:</b></td>
			    <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="ShipCompany" size="25" value="<%= sShipCustCompany %>"></td>
			    <td nowrap align="right">&nbsp;</td>
			    <td nowrap>&nbsp;</td>
			  </tr>
			  <tr>
			    <td nowrap align="right" height="15"></td>
			    <td nowrap height="15"></td>
			    <td nowrap align="right" height="15"></td>
			    <td nowrap height="15"></td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>Street Address:</b></td>
			    <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="ShipAddress1" size="25" value="<%= sShipCustAddress1%>"></td>
			    <td nowrap align="right"><b>Street Address 2:</b></td>
			    <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="50" name="ShipAddress2" size="25" value="<%= sShipCustAddress2 %>"></td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>City:</b></td>
			    <td nowrap><input type="text" maxlength="50" name="ShipCity" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustCity %>"></td>
			    <td nowrap align="right"><b>State/Province:</b></td>
			    <td nowrap><select size="1" name="ShipState" style="<%= sFormDesign %>"><option value="<%= sShipCustState %>"><%= sShipCustStateName %></option><%= sStateList %></select></td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>Zip/Postal
                Code:</b></td>
			    <td nowrap><input type="text" maxlength="12" name="ShipZip" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustZip %>"></td>
			    <td nowrap align="right"><b>Country:</b></td>
			    <td nowrap><select size="1" name="ShipCountry" Style="<%= sFormDesign%>"><option value="<%= sShipCustCountry %>"><%= sShipCustCountryName %></option><%= shipCtryList %></select></td>
			  </tr>
			  <tr>
			    <td nowrap align="right" height="15"></td>
			    <td nowrap height="15"></td>
			    <td nowrap align="right" height="15"></td>
			    <td nowrap height="15"></td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>Phone:</b></td>
			    <td nowrap><input type="text" maxlength="20" name="ShipPhone" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustPhone %>"></td>
			    <td nowrap align="right"><b>E-Mail:</b></td>
			    <td nowrap><input type="text" Style="<%= sFormDesign%>" maxlength="100" name="ShipEmail" size="25" value="<%= sShipCustEmail %>">
                </td>
			  </tr>
			  <tr>
			    <td nowrap align="right"><b>Fax:</b></td>
			    <td nowrap><input type="text" maxlength="20" name="ShipFax" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustFax %>"></td>
			    <td nowrap align="right">&nbsp;</td>
			    <td nowrap>&nbsp;
                </td>
			  </tr>
			  <tr>
			    <td nowrap align="right"></td>
			    <td nowrap></td>
			    <td nowrap align="right"></td>
			    <td nowrap></td>
			  </tr>   
			  <tr>
			    <td width="100%" colspan="4"><center><img src="<%= C_BTN23 %>" onClick="javascript:clearShipping(form1);"></center><br></td>
			  </tr>
			</table>
          </td></tr>    
<% if iShip <> 0 Then 
       If sShipMethods <> "" Then %> 
        <tr><td width="100%" class="tdContent2">  
		    <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%">
		      <tr>
		        <td width="100%">

		          <table border="0" width="100%" cellspacing="0" cellpadding="2">
		            <tr>
		              <td colspan="2" width="100%" class="tdContentBar"> Shipping Method<span lang="en-us">
                      <font size="1">(for United Kingdom, select UPS WorldWide 
                      Expedited.&nbsp; we do ship to APO, PPO and FPO addresses)</font></span></td>
		            </tr>
		          </table>
		        </td>
		      </tr>    
		      <tr>
		        <td align="center" colspan="2"><br><select name="Shipping" style="<%= sFormDesign %>"><%= sShipMethods%></select></td>
		        </tr>
		        <tr><td>&nbsp;</td></tr> 
		      </table>
	      <%  End If 
     End If %>
          <!-- Payment Method -->
           <tr><td width="100%" colspan="2">
              <table border="0" width="100%" cellspacing="0" cellpadding="2" class="tdContent2">
                <tr>
                  <td width="100%" class="tdContentBar"> Select Payment Method</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr><td width="100%" colspan="2">
              <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
	 	        <tr>
	 	          <td height="50" valign="center" nowrap align="center">
		            <select style="<%= C_FORMDESIGN %>" name="PaymentMethod"><%= sPaymentList %></select>	   
		          </td>
		        </tr>		
	          </table>	
	        </td>
	      </tr>
	       <tr><td width="100%" class="tdContent2">  
		      <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%">
		        <tr>
		          <td width="100%">
		            <table border="0" width="100%" cellspacing="0" cellpadding="2">
		              <tr>
		                <td colspan="2" width="100%" class="tdContentBar">Special Instructions<span lang="en-us">
                        <font size="1">(e.g., if you need order by a certain 
                        date)</font></span></td>
		              </tr>
		            </table>
		          </td></tr>
		        <tr> 
		          <td width="100%" class="tdContent2" align="center">
		            <br>
		            <textarea rows="4" name="Instructions" cols="60" style="<%= sFormDesign%>"></textarea>  
		           </td>
		        </tr>
		        <tr><td>&nbsp;</td></tr>
		      </table>
	        </td></tr>
	      <% If NOT (bLoggedIn) Then %>
	      <tr><td width="100%" class="tdContent2">  
		      <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%">
		        <tr><td width="100%">
		            <table border="0" width="100%" cellspacing="0" cellpadding="2">
		              <tr>
		                <td colspan="2" width="100%" class="tdContentBar"> New Customers: Choose Password </td>
		              </tr>
		            </table>
		          </td></tr>
		        <tr>
		          <td class="tdContent2">
                    In order to serve you better, an account will be created
            for you as part of the checkout process. This will facilitate a
            speedier checkout for future orders. 
To specify a password, please enter it below. Otherwise, a password will be generated for you. 
			        <center>
		            <table border="0" width="100%" class="tdContent2" align="center">
					    <tr><td colspan="2">&nbsp;</td></tr>
					    <tr>
					      <td width="50%" align="right"><b>Password:</b></td>
					      <td width="50%">
                          <input type="password" name="Password" maxlength="10" title="Password" style="<%= sFormDesign%>" size="20"></td>
					    </tr>
					    <tr>
					      <td width="50%" align="right"><b>Password Confirmation:</b></td>
					      <td width="50%">
                          <input type="password" name="Password2" maxlength="10" title="Password Confirmation" style="<%= sFormDesign%>" size="20"></td>
					    </tr>
					    <tr><td colspan="2">&nbsp;</td></tr>
		            </table>
			        </center>	    
                  </td>
                </tr>
              </table>   
            </td>    
          </tr>   
	      <% 
	End If 
	closeObj(cnn)
	%>
	      <tr>
	        <td width="100%" class="tdContentBar"> Click
        &quot;Continue&quot; to Verify Charges and Enter Payment Information</td>
	      </tr>
	      <tr align="center">
	        <td><Br>
	          <input type="image" src="<%= C_BTN20%>" name="Verify" border="0">
		<br>

	          </td>
	          </tr>
    </table>

				 <input type="hidden" name="FreeShip" value="<%= blnFree %>">
				 <input type="hidden" name="TotalPrice" value="<%= sTotalPrice %>">
				  <input type="hidden" name="bShip" value="<%= iShip %>">
				</form>