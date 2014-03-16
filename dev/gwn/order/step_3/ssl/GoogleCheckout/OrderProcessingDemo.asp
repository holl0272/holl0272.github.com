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

'*******************************************************************************
' Please refer to the Google Checkout ASP Sample Code Documentation
' for requirements and guidelines on how to use the sample code.
'
' This sample code demonstrates how to utilize the Google Checkout ASP 
' Client Library to generate an order processing request and transmit it 
' to Google Checkout.
'
' To use this code, you need to set your Merchant ID (strMerchantId) and 
' Merchant Key (strMerchantKey) in the GlobalAPIFunctions.asp file. 
' Your Merchant ID and Merchant Key can be found in your Google Checkout 
' Account page under the Settings tab.
'*******************************************************************************

' Include ASP libraries used in order processing demo
%>

<!--#INCLUDE file="GlobalAPIFunctions.asp"-->
<!--#INCLUDE file="OrderProcessingAPIFunctions.asp"-->
<!--#INCLUDE file="ResponseHandlerAPIFunctions.asp"-->

<%

'**************** Build Order Processing Commands *******************

' Define objects used to create and send Order Processing API requests
Dim xmlRequest
Dim xmlResponse

Dim attrGoogleOrderNumber
Dim elemAmount 
Dim elemReason
Dim elemComment
Dim elemCarrier
Dim elemTrackingNumber
Dim elemMessage
Dim elemSendEmail
Dim elemMerchantOrderNumber

Dim transmitResponse

' This section creates a <charge-order> request
' Comment out the following section and uncomment another section 
' to create another request
attrGoogleOrderNumber = "841171949013218"
elemAmount = "100.00"
xmlRequest = createChargeOrder(attrGoogleOrderNumber, elemAmount)

' This section creates a <refund-order> request
' attrGoogleOrderNumber = "841171949013218"
' elemReason = "Buyer requested refund."
' elemAmount = "120.00"
' elemComment = "Buyer is not happy with the product."
' xmlRequest = createRefundOrder(attrGoogleOrderNumber, elemReason, _
'     elemAmount, elemComment)

' This section creates a <cancel-order> request
' attrGoogleOrderNumber = "841171949013218"
' elemReason = "Buyer cancelled the order."
' elemComment = "Buyer found a better deal."
' xmlRequest = createCancelOrder(attrGoogleOrderNumber, elemReason, elemComment)

' This section creates a <process-order> request
' attrGoogleOrderNumber = "841171949013218"
' xmlRequest = createProcessOrder(attrGoogleOrderNumber)

' This section creates a <deliver-order> request
' attrGoogleOrderNumber = "841171949013218"
' elemCarrier = "UPS"
' elemTrackingNumber = "Z5498W45987123684"
' xmlRequest = createDeliverOrder(attrGoogleOrderNumber, elemCarrier, _
'    elemTrackingNumber)

' This section creates an <add-tracking-data> request
' attrGoogleOrderNumber = "841171949013218"
' elemCarrier = "UPS"
' elemTrackingNumber = "Z9842W69871281267"
' xmlRequest = createAddTrackingData(attrGoogleOrderNumber, elemCarrier, _
'    elemTrackingNumber)

' This section creates an <archive-order> request
' attrGoogleOrderNumber = "841171949013218"
' xmlRequest = createArchiveOrder(attrGoogleOrderNumber)

' This section creates an <unarchive-order> request
' attrGoogleOrderNumber = "841171949013218"
' xmlRequest = createUnarchiveOrder(attrGoogleOrderNumber)

' This section creates a <send-buyer-message> request
' attrGoogleOrderNumber = "841171949013218"
' elemMessage = "Dear Customer, due to a high volume of orders, your order " _
'     & "will not be charged and shipped until next week."
' elemSendEmail = "true"
' xmlRequest = createSendBuyerMessage(attrGoogleOrderNumber, elemMessage, _
'    elemSendEmail)

' This section creates an <add-merchant-order-number> request
' attrGoogleOrderNumber = "841171949013218"
' elemMerchantOrderNumber = "MyOrderNumber012345"
' xmlRequest = createAddMerchantOrderNumber(attrGoogleOrderNumber, _
'    elemMerchantOrderNumber)

' The following HTML page calls the displayDiagnoseResponse function,
' which is defined in GlobalAPIFunctions.asp, to verify that the 
' API request contains valid XML. If the request does contain valid XML,
' the page sends the request to Google Checkout by calling the sendRequest 
' function, which is also defined in GlobalAPIFunctions.asp.
' The page then calls the processXmlData function, which is defined in
' ResponseHandlerAPIFunctions.asp, to handle Google Checkout's API response.
'
' If the request does not contain valid XML, you will see a link to 
' a tool that lets you edit and recheck the XML. The code for that 
' tool is in the <b>DebuggingTool.asp</b> file, which is also 
' included in the <b>checkout-asp-samplecode.zip</b> file.
%>

<html>
<head>
    <style type="text/css">@import url(googleCheckout.css);</style>
</head>
<body>
<p style="text-align:center">
<table class="table-1" cellspacing="5" cellpadding="5">
    <tr><td style="padding-bottom:20px;text-align:center"><h2>
        Order Processing Command
    </h2></td></tr>
    <tr><td style="padding-bottom:20px">
        <p><b>Order Processing Command XML:</b></p>
        <p><%=Server.HTMLEncode(xmlRequest)%></p>
    </td></tr>
<%
    ' Validate Request XML
    DisplayDiagnoseResponse xmlRequest, requestDiagnoseUrl, _
        xmlRequest, "diagnose"

    Response.write "<tr><td style=""padding-bottom:20px"">" & _
        "<p><b>Synchronous Response Received:</b></p>"

    ' Send the request and receive a response
    transmitResponse = SendRequest(xmlRequest, requestUrl)

    ' Process the response
    Response.write "<p>" & ProcessXmlData(transmitResponse) & "</p></td></tr>"

%>
</table>
</p>
</body>
</html>