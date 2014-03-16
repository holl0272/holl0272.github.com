<%
'*******************************************************************************
' Copyright (C) 2006 Google Inc.
'  
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'  
'      http://www.apache.org/licenses/LICENSE-2.0
'  
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'*******************************************************************************
%>

<!--#INCLUDE file="GlobalAPIFunctions.asp"--> 
<!--#INCLUDE file="ResponseHandlerAPIFunctions.asp"--> 

<%
Dim transmitResponse
Dim diagnoseResponse
Dim bValidated
Dim xml
Dim b64signature
Dim b64cart
Dim checkoutPostData


If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    ' Get <checkout-shopping-cart> XML
    xml = Trim(Request.Form("xml"))

    If Request.Form("toolType") = "Submit Cart to Google Checkout" Then

        transmitResponse = SendRequest(xml, requestUrl)
        ProcessXmlData(transmitResponse)
        Response.End

    ElseIf Request.Form("toolType") = "Display HTML Form for Checkout" Then

        ' Use the cart XML and your Merchant Key to calculate the HMAC-SHA1 
        ' value and Base64-encode the Cart XML and the signature before posting
        b64cart = cryptObj.base64Encode(xml)
        b64signature = cryptObj.generateSignature(xml,strMerchantKey)

        checkoutPostData = "cart=" & Server.urlencode(b64cart) & _
            "&signature=" & Server.urlencode(b64signature)

        ' Free object
        Set cryptObj = Nothing

        ' Log <checkout-shopping-cart> XML
        LogMessage logFilename, checkoutPostData

' The following HTML page displays some information about the POST request
' that will be submitted to Google Checkout if you click the Google Checkout 
' button that appears on the page. The Google Checkout button is embedded in 
' a form similar to the form you want to include on your site. The form sends
' the request to Google Checkout and shows you an interface similar to what 
' your customer would see after clicking the Google Checkout button.
'
' Note: This page also calls the displayDiagnoseResponse function,
' which is defined in GlobalAPIFunctions., to verify that the 
' API request contains valid XML. If the request does not contain
' valid XML, you will see a link to a tool that lets you edit and
' recheck the XML. The code for that tool is in the 
' <b>DebuggingTool.asp</b> file, which is also included
' in the <b>ClassicASP.zip</b> file.
%>
<html>
<head>
    <style type="text/css">@import url(gbuy.css);</style>
</head>
<body>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
    <tr><td style="padding-bottom:20px;text-align:center"><h2>
        Place a New Order
    </h2></td></tr>

    <!-- Print the shopping cart XML -->
    <tr><td style="padding-bottom:20px">
        <p><b>This is the &lt;checkout-shopping-cart&gt; XML before it is 
        base64-encoded:</b></p>
        <p><%=Server.HTMLEncode(xml)%></p>
    </td></tr>

    <!-- Print the HMAC-SHA1 signature -->
    <tr><td style="padding-bottom:20px">
        <p><b>This is the base64-encoded signature:</b></p>
        <p><%=Server.HTMLEncode(b64signature)%></p>
    </td></tr>

    <!-- Print Error message if the cart XML is invalid -->
<%
        displayDiagnoseResponse checkoutPostData, checkoutDiagnoseUrl, _
            xml, "debug"
%>
    <tr><td style="padding-bottom:20px">
        <p><b>Click checkout to post this cart.</b></p>
        <p><form method="POST" action="<%=checkoutUrl%>">
            <input type="hidden" name="cart" value="<%=b64cart%>">
            <input type="hidden" name="signature" value="<%=b64signature%>">
            <input type="image" name="Checkout" alt="Checkout" 
            src="http://checkout.google.com/buttons/checkout.gif?
            merchant_id=<%=strMerchantId%>&
            w=180&h=46&style=white&variant=text&loc=en_US" 
            height="46" width="180">
        </form></p>
    </td></tr>
    </table>
    </p>
</body>
</html>
<%
        Response.End
    End If
Else
    xml = ""
End If
%>
<html>
<head>
    <style type="text/css">@import url(gbuy.css);</style>
</head>
<body>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
    <tr><td width="100%" style="text-align:center">
        <h2>Google Checkout API XML Debugging Tool</h2>
    </td></tr>
    <tr><td style="text-align:left">
        <form method="POST" 
        action="<%=Request.ServerVariables("REQUEST_URI")%>">
        <p><b>XML:</b></p>
        <p><textarea name="xml" cols="80" rows="20"><%=xml%></textarea></p>
        <p><table style="text-align:left" cellspacing="5" cellpadding="5">
            <tr><td><input name="toolType" type="submit" 
            value="Validate XML"></td>
            <td><input name="toolType" type="submit" 
            value="Display HTML Form for Checkout"></td>
            </tr><tr>
            <td><input name="toolType" type="submit" 
            value="Send Order Processing Command"></td>
            <td><input name="toolType" type="submit" 
            value="Submit Cart to Google Checkout"></td>
            </tr>
        </table></p>
        </form>
        </td></tr>
    </table>
    </p>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
<%    
If Request.ServerVariables("REQUEST_METHOD") = "POST" And _
    Request.Form("toolType") = "Validate XML" Then

    ' Print Error message if the XML is invalid
    bValidated = displayDiagnoseResponse(xml, requestDiagnoseUrl, _
        xml, "diagnose")

    If bValidated = true Then
%>
        <tr><td colspan="2">
            <span style="text-align:center;color:green">
            <h2>This XML is Validated!</h2>
            </span>
        </td></tr>
<% 
    End If

    Response.write "</table>"

ElseIf Request.ServerVariables("REQUEST_METHOD") = "POST" And _
    Request.Form("toolType") = "Send Order Processing Command" Then
%>
    <table class="table-1" cellspacing="5" cellpadding="5">
        <tr><td style="padding-bottom:20px;text-align:center"><h2>
        Order Processing Command
        </h2></td></tr>
        <tr><td style="padding-bottom:20px">
            <p><b>Order Processing Command XML:</b></p>
            <p><%=Server.HTMLEncode(xml)%></p>
        </td></tr>
<%
    ' Validate Request XML
    displayDiagnoseResponse xml, requestDiagnoseUrl, xml, "diagnose"

    Response.write "<tr><td style=""padding-bottom:20px"">" & _
        "<p><b>Synchronous Response Received:</b></p>"

    ' Send the request and receive a response
    transmitResponse = SendRequest(xml, requestUrl)

    ' Process the response
    Response.write "<p>" & ProcessXmlData(transmitResponse) & "</p></td></tr>"
    Response.write "</table>"

End If
%>
</p>
</body>
</html>