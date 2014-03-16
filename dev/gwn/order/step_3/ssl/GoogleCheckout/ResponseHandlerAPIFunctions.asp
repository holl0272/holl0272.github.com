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
' "ResponseHandlerAPIFunctions.asp" is a sample set of functions for recognizing
' response XML messages and calling the appropriate function for 
' further processing.
' You should modify the functions that handle the responses to suit your needs.
'
' If you are implementing the Notification and Order Processing APIs,
' you should also include the OrderProcessingAPIFunctions file, which
' contains functions that handle Order Processing API requests. 
'*******************************************************************************
%>

<!--#INCLUDE file="NotificationAPIFunctions.asp"--> 
<!--#INCLUDE file="MerchantCalculationsAPIFunctions.asp"--> 

<%
'*******************************************************************************
' The processXmlData function creates a DOM object representation of the
' XML document received from Google Checkout. It then evaluates the root tag 
' of the document to determine which function should handle the document.
'
' This function routes the XML responses that Google Checkout sends in response 
' to API requests. These replies are sent to one of the other three functions
' in this library.
'
' This function also routes Merchant Calculations API requests and
' Notification API requests. Those requests are processed by functions
' in the MerchantCalculationsAPIFunctions.asp and
' NotificationAPIFunctions.asp libraries, respectively.
'
' Input:   xmlData  The XML document sent by the Google Checkout server.
'*******************************************************************************
Function processXmlData(xmlData)

    Dim domResponseObj

    Set domResponseObj = Server.CreateObject(strMsxmlDomDocument)
    domResponseObj.loadXml xmlData

    Dim messageRecognizer
    messageRecognizer = domResponseObj.documentElement.tagName

    ' Select the appropriate function to handle the XML document
    ' by evaluating the root tag of the document. Functions to
    ' handle the following types of responses are contained in
    ' this document:
    '     <request-received>
    '     <error>
    '     <diagnosis>
    '     <checkout-redirect>
    '
    ' This function routes the following types of responses
    ' to the MerchantCalculationsAPIFunctions.asp file:
    '     <merchant-calculation-callback>
    '
    ' This function routes the following types of responses
    ' to the NotificationAPIFunctions.asp file:
    '     <new-order-notification>
    '     <order-state-change-notification>
    '     <charge-amount-notification>
    '     <chargeback-amount-notification>
    '     <refund-amount-notification>
    '     <risk-information-notification>

    Select Case messageRecognizer

        ' <request-received> received
        Case "request-received"
            processRequestReceivedResponse domResponseObj
         
        ' <error> received
        Case "error"
            processErrorResponse domResponseObj

        ' <diagnosis> received
        Case "diagnosis"
            processDiagnosisResponse domResponseObj

        ' <checkout-redirect> received
        Case "checkout-redirect"
            processCheckoutRedirect domResponseObj

        ' +++ CHANGE ME +++
        ' The following case is only for partners who are implementing 
        ' the Merchant Calculations API. If you are not implementing
        ' the Merchant Calculations API, you may ignore this case.
        ' <merchant-calculation-callback> received
        Case "merchant-calculation-callback"
            processMerchantCalculationCallback domResponseObj


        ' +++ CHANGE ME +++
        ' The following cases are only for partners who are
        ' implementing the Notification API. If you are not
        ' implementing the Notification API, you may ignore
        ' the remaining cases in this function.

        ' <new-order-notification> received
        Case "new-order-notification"
            processNewOrderNotification domResponseObj
     
        ' <order-state-change-notification> received
        Case "order-state-change-notification"
            domProcessOrderStateChangeNotification domResponseObj
     
        ' <charge-amount-notification> received
        Case "charge-amount-notification"
            processChargeAmountNotification domResponseObj
         
        ' <chargeback-amount-notification> received
        Case "chargeback-amount-notification"
            processChargebackAmountNotification domResponseObj
         
        ' <refund-amount-notification> received
        Case "refund-amount-notification"
            processRefundAmountNotification domResponseObj
         
        ' <risk-information-notification> received
        Case "risk-information-notification"
            processRiskInformationNotification domResponseObj
         
        ' None of the above: message is not recognized.
        ' You should not remove this case.
        Case Else

    End Select 

End Function


'********* Functions for processing synchronous response messages ********

'*******************************************************************************
' The processRequestReceivedResponse function receives a synchronous
' Google Checkout response to an API request originating from your site. This
' function indicates that your API request contained properly formed
' XML but does not indicate whether your request was processed successfully.
'
' Input:  domResponseObj    <request-received> XML DOM
'*******************************************************************************
Function processRequestReceivedResponse(domResponseObj)
    ' +++ CHANGE ME +++
    ' You may need to modify this function if you wish to log information
    ' or perform other actions when you receive a Google Checkout
    ' <request-received> response. The <request-received> response indicates
    ' that you sent a properly formed XML request to Google Checkout.  However, 
    ' this response does not indicate whether your request was processed 
    ' successfully.
    Response.write Server.HTMLEncode(domResponseObj.xml)
End Function


'*******************************************************************************
' The processErrorResponse function receives a synchronous Google Checkout 
' response to an API request originating from your site. This function indicates
' that your API request was not processed. A request might not be processed
' if it does not contain properly formed XML or if it does not contain a
' valid merchant ID and merchant key.
'
' Input:  domResponseObj    <error> XML DOM
'*******************************************************************************
Function processErrorResponse(domResponseObj)
    ' +++ CHANGE ME +++
    ' You may need to modify this function if you wish to log
    ' information or perform other actions when you receive
    ' a Google Checkout <error> response. The <error> response indicates that
    ' you sent an invalid XML request to Google Checkout and contains
    ' information explaining why the request was invalid.
    Response.write domResponseObj.xml
End Function


'*******************************************************************************
' The ProcessDiagnosisResponse function receives a synchronous Google
' Checkout response to an API request that was sent to the Google Checkout
' XML validator. You can submit a request to the validator by appending
' the text "/diagnose" to the POST target URL. The response to a
' diagnostic request contains a list of any warnings returned by
' the Google Checkout validator.
'
' Input:  domResponseObj    <diagnosis> XML DOM
'*******************************************************************************
Function processDiagnosisResponse(domResponseObj)
    ' +++ CHANGE ME +++
    ' You may need to modify this function if you wish to log
    ' warnings or perform other actions when you receive
    ' a Google Checkout <diagnosis> response. The <diagnosis> response contains
    ' warnings that the Google Checkout XML validator generated when
    ' evaluating your XML request.
    Response.write "<i>Diagnosis response message received:</i><br>"
    Response.write Server.HTMLEncode(domResponseObj.xml)
End Function

'*******************************************************************************
' Process <checkout-redirect> synchronous response message
'
' Input:  domResponseObj    <checkout-redirect> XML DOM
'*******************************************************************************
Function processCheckoutRedirect(domResponseObj)

    ' Define objects used to process <checkout-redirect> response
    Dim domResponseObjRoot
    Dim redirectUrlList
    Dim strRedirectUrl

    ' Identify the URL to which the customer should be redirected
    Set domResponseObjRoot = domResponseObj.documentElement
    Set redirectUrlList = _
        domResponseObjRoot.getElementsByTagname("redirect-url")
    strRedirectUrl = redirectUrlList(0).text

    ' Redirect the customer to the URL
    Response.redirect strRedirectUrl

    ' Release objects used to process <checkout-redirect> response
    Set domResponseObjRoot = Nothing
    Set redirectUrlList = Nothing

End Function

%>