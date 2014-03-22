<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.5

'@FILENAME: neworder.asp
 

'@DESCRIPTION: Retrieves Order

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

' #326 - MS
' #308 - MS

vDebug = 0
If vDebug = 1 Then Response.Buffer = True

'@BEGINCODE
Dim sSql, rsAllOrders, sProdID, aProduct, sProdName, sProdPrice, iProdAttrNum, iCounter, sCondition, iAttCounter
Dim sAttrUnitPrice, dUnitPrice, iQuantity, dProductSubtotal, dTotalPrice, iOrderID, aProdAttrID, sEmail, sPassword
Dim iProductCounter, sBgColor, sFontFace, sFontColor, iFontSize,sBkGrnd, rsProdOrders, rsOrderProdAtt
Dim rsAttributeDetails,rsAttribute ,sAmount

Dim mstrCallingPage

	' Determine if it is recalculate action
	' Product counter initialize
	iProductCounter = 0	
	dTotalPrice = 0 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Retrieve Order Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="keywords">
<meta name="description" content="description">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">

<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">

<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<% writeCurrencyConverterOpeningScript %>
</head>

<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<% 	Call ShowMyAccountBreadCrumbsTrail("", False) %>
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
  <tr>
    <td>
      <table width="95%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="center" class="tdMiddleTopBanner">Retrieve Your Order</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner">Your
            previous order is shown below. To add items to your current order,
            select 'Add to Cart'.
			</td>
          </tr>
          <tr>
            <td class="tdContent2">        
              <table border="0" width="100%" cellspacing="0" cellpadding="5">
   <%
	'-----------------------------------------------------------------
	' Collect all orders associated with Old Order ::: Begin
	'-----------------------------------------------------------------
	' Check cookies and other indicators of login
	If Not isLoggedIn Then
		mstrCallingPage = LoadRequestValue("PrevPage")
		If Len(mstrCallingPage) = 0 Then mstrCallingPage = Request.ServerVariables("SCRIPT_NAME")
		Call ProtectThisPage(mstrCallingPage)
	Else
			' #326 added orderIsComplete = 1
			sSql = "SELECT * FROM sfOrders WHERE orderIsComplete = 1 and ordercustID = " & visitorLoggedInCustomerID & " Order by orderID Desc"
			If vDebug = 1 Then Response.Write "<br /> " & sSql
			Set rsAllOrders = cnn.execute(sSql)

			If rsAllOrders.EOF OR rsAllOrders.BOF Then 
	%>
			    <tr>
			      <td colspan="4" width="40%" align="center">		
		        	  <table border="0" width="50%" cellpadding="0" cellspacing="1">
		        	  	  <tr>
						       <td width="100%" height="50" valign="middle" align="center" class="tdContent"><b>No Previous Orders Found</b></td>		        
						  </tr>
		              </table>	
		           
			      </td>
			    </tr> 		          		
	<%	Else %>
		<!--webbot bot="PurpleText" PREVIEW="Begin Optional Confirmation Message Display" -->
		<tr>
			<td valign="bottom" class="tdContent2" colspan="5">
			<% Call WriteThankYouMessage %>
			</td>
		</tr>		
		<!--webbot bot="PurpleText" PREVIEW="End Optional Confirmation Message Display" -->
	<%		
			Dim sCheckProduct
			Dim sTmpOrderID
			Dim pstrAttributeDetails

			Do While Not rsAllOrders.EOF
				sTmpOrderID = rsAllOrders.Fields("orderID")
	%>	
				<tr>
					<td valign="bottom" class="tdContent2" colspan="5"><b>Order ID: <%= sTmpOrderID %></b></td>
				</tr>		
				<tr>
					<td width="40%" class="tdContentBar">item</td>
					<td width="15%" align="center" class="tdContentBar">unit price</td>
					<td width="15%" align="center" class="tdContentBar">qty</td>
					<td width="15%" align="center" class="tdContentBar">price</td>
				</tr>
	<%	
				sSql = "SELECT * FROM sfOrderDetails WHERE odrdtOrderId = " & makeInputSafe(sTmpOrderID) & " Order by odrdtOrderId"
				Set rsProdOrders = cnn.Execute(sSQL)				
				Do While NOT rsProdOrders.EOF
					iOrderID = rsProdOrders.Fields("odrdtID")
					sProdID = rsProdOrders.Fields("odrdtProductID")
					iQuantity = rsProdOrders.Fields("odrdtQuantity")
					sProdName = rsProdOrders.Fields("odrdtProductName")
					sProdPrice = rsProdOrders.Fields("odrdtPrice")
	    	    	'Get an array of 3 values from getProduct()
				   	'++ On Error Resume Next
					ReDim aProduct(3)
					aProduct = getProduct(sProdID)		

					sCheckProduct = aProduct(0)

	  				iProdAttrNum = aProduct(2)	  			
					If Trim(sCheckProduct) = "" Then
						sCheckProduct = "deleted"
					End If

					'If not an array, then the product does not exist 
					If NOT IsArray(aProduct) Then
						Response.Write "<br />Product No Longer In Inventory"
						'++ Needs to MoveNext to iterate through the rest of the order			
					Else
						If NOT IsNumeric(iProdAttrNum)Then 
							iProdAttrNum = 0
						End If	
						
						' Response Write all Output
						If vDebug = 1 And IsArray(aProdAttrID) Then 
							Response.Write "<p>Product = " & sProdID & "<br />ProdName = " & sProdName & "<br />ProdPrice = " & sProdPrice & "<br />ProdAttrNum = " & iProdAttrNum
						
							For iCounter = 0 To iProdAttrNum -1 
								Response.Write "<br />Attribute :" & aProdAttrID(iCounter)
							Next			
					
						End If	 
				
						iProductCounter = iProductCounter + 1
		
						dim fontclass
						' Do alternating colors and fonts	
						If (iProductCounter mod 2) = 1 Then 
							fontclass = "tdAltFont1"
						Else 	
							fontclass = "tdAltFont2"
						End If	
		
						'----------------------------------------------------
						'Get Order Attributes
						'----------------------------------------------------
						pstrAttributeDetails = ""
						sSQL = "SELECT odrattrName, odrattrAttribute FROM sfOrderAttributes WHERE odrattrOrderDetailId = " & rsProdOrders.Fields("odrdtID")
						Set rsOrderProdAtt = cnn.execute(sSql)
						Do While Not rsOrderProdAtt.EOF
							pstrAttributeDetails = rsOrderProdAtt.Fields("odrattrName") & ": " & rsOrderProdAtt.Fields("odrattrAttribute") & "<br />" & vbcrlf
							rsOrderProdAtt.MoveNext
						Loop
					%>
<form name="<%= sProdID %>" action="addproduct.asp" method="post">
	                <tr>
	                  <td width="40%" valign="top" class='<%= fontClass %>' nowrap>
	                  <b><%= sProdName %></b><br /><%= pstrAttributeDetails %>
	                  </td>
	<%
	' Set Unit Price for Product

		dUnitPrice = cdbl(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
		dProductSubtotal = iQuantity * (cDbl(sAttrUnitPrice) + cDbl(sProdPrice))

		dTotalPrice = dTotalPrice + cDbl(dProductSubtotal)
%>
					  <td width="15%" align="center" class='<%= fontClass %>' valign="top" nowrap><% writeCustomCurrency(dUnitPrice) %></td>
					  <td width="15%" align="center" class='<%= fontClass %>' valign="top" nowrap><input type="text" class="formDesign" name="QUANTITY" size="2" value="<%= iQuantity %>"></td>
					  <td width="15%" align="center" class='<%= fontClass %>' valign="top" nowrap><% writeCustomCurrency(dProductSubtotal) %></td>	          
	                </tr>
<%	

	'--------------------------------------------------------------------
	'End Get Order Attributes
	'--------------------------------------------------------------------
	'--------------------------------------------------------------------
	'Get Product Attributes
	'--------------------------------------------------------------------
	
	If sCheckProduct = "deleted" Then
	%>
	<tr>
	  <td><font class="Content_Small"><b>This Product is No Longer Available</b></font></td>
	</tr>
	<%
	Else
		Dim iCount
		dim bSelected
		Dim pstrAttributeName
	
		sAttrUnitPrice = 0

		' Iterate Through Attributes
		iAttCounter = 1
		If iProdAttrNum > 0 Then
			Set rsAttribute = CreateObject("ADODB.RecordSet")
			Set rsAttributeDetails = CreateObject("ADODB.RecordSet")
			rsAttribute.CursorLocation = adUseClient
			rsAttributeDetails.CursorLocation = adUseClient
			sSQL = "SELECT * FROM sfAttributes WHERE attrProdId ='" & makeInputSafe(sProdId) & "'"
			rsAttribute.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			While Not rsAttribute.EOF 								
%>                    
<tr>
	<td><font class="Content_Small"><%= rsAttribute("attrName") %>:</font></td>
	<td>
	<select size="1" name="attr<%= iAttCounter %>" class="formDesign">
<%
		sSQL = "SELECT * FROM sfAttributeDetail WHERE attrdtAttributeId = " & rsAttribute("attrID")
		rsAttributeDetails.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	   '#789 DJP
		If cblnSQLDatabase Then
	   		sSQL = "SELECT odrattrName, odrattrAttribute FROM sfOrderAttributes WHERE odrattrOrderDetailId = " & rsProdOrders.Fields("odrdtID") & " AND odrattrName Like '" & rsAttribute("attrName") & "'"
		Else
	   		sSQL = "SELECT odrattrName, odrattrAttribute FROM sfOrderAttributes WHERE odrattrOrderDetailId = " & rsProdOrders.Fields("odrdtID") & " AND odrattrName = '" & rsAttribute("attrName") & "'"
		End If
	    Set rsOrderProdAtt = cnn.execute(sSql)
	
		Do While Not rsAttributeDetails.EOF 
			sAmount = ""
			Select Case rsAttributeDetails("attrdtType")
				Case 1 
					sAmount = " (add " & FormatCurrency(rsAttributeDetails("attrdtPrice")) & ")"
				Case 2 
					sAmount = " (subtract " & FormatCurrency(rsAttributeDetails("attrdtPrice")) & ")"
			End Select
		
			If Not rsOrderProdAtt.EOF Then
				pstrAttributeName = Trim(rsOrderProdAtt.Fields("odrattrAttribute").Value & "")
				pstrAttributeName = Trim(Replace(pstrAttributeName, rsOrderProdAtt.Fields("odrattrName").Value, "", 1, 1))
				If Left(pstrAttributeName, 2) = ": " Then pstrAttributeName = Right(pstrAttributeName, Len(pstrAttributeName) - 2)
			End If

			If pstrAttributeName = Trim(rsAttributeDetails.Fields("attrdtName").Value) then
			%><option selected value="<%= rsAttributeDetails("attrdtID") %>"><%= rsAttributeDetails("attrdtName") & sAmount %></option><%
			Else
			%><option value="<%= rsAttributeDetails("attrdtID") %>"><%= rsAttributeDetails("attrdtName") & sAmount %></option><%
			End IF
		rsAttributeDetails.MoveNext 
	Loop 
	rsAttributeDetails.Close 
%>
	</select><br />
	</td>
	</tr>
<%  
	iAttCounter = iAttCounter + 1
	rsAttribute.MoveNext 
	Wend
	rsAttribute.Close 
					
	Set rsAttribute = Nothing
	Set rsAttributeDetails = Nothing	
	rsOrderProdAtt.Close 
	Set rsOrderProdAtt = Nothing
End If 
%>
<tr><td colspan=2><%

	If cblnSF5AE Then
		SearchResults_GetProductInventory sProdID
		SearchResults_ShowMTPricesLink sProdID
	End If
	%></td></tr>

<tr><td colspan=2><% SearchResults_GetGiftWrap sProdID %></td></tr>

<tr>
<td colspan="3">&nbsp;</td>
<td width="15%" align="center" valign="top">
<input type="hidden" name="PRODUCT_ID" value="<%=sProdID%>">
<input type="image" class="inputImage" src="<%= C_BTN03 %>" name="AddProduct">
</td>
</tr>
</form>
<%
	End If
	' End IsArray If
	End If
	
	' Move to next RecordSet
	rsProdOrders.MoveNext		
	' loop through recordset	
	Loop
%>
<tr>
<td colspan="5" width="100%">
<hr size="2" width="100%"> 
</td>
<%
	rsAllOrders.MoveNext
	'Loop thorugh next order
	Loop
	' End if not empty orders if
	End If
	
	'-----------------------------------------------------------
	' END PRODUCT DETAIL OUTPUT --------------------------------
	'-----------------------------------------------------------
   ' End rsAllOrders If
	End If %>
	</tr>  
</table>	      
            </td>
          </tr>
        </table>

                  	</td>
                		</tr>
                                </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
</body>
</html>
    <%
closeObj(rsAllOrders)
Call cleanup_dbconnopen
%>