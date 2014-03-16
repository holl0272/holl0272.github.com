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
' "OrderProcessingAPIFunctions.asp" is a client library of functions that 
' enables you to systematically generate XML for Order Processing API requests.
' 
' You should also look at the Demo files to learn more about how to call
' each function and what it returns.
'*******************************************************************************


'******************* Order Processing API ********************

'*******************************************************************************
' The createArchiveOrder function is a wrapper function that calls the
' changeOrderState function. The changeOrderState function, in turn,
' creates an <archive-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). 
' The <archive-order> command moves an order from the Inbox in the 
' Google Checkout Merchant Center to the Archive folder.
'
' Input:    attrGoogleOrderNumber       A number, assigned by Google Checkout,
'                                       that uniquely identifies an order.
' Returns:  <archive-order> XML
'*******************************************************************************
Function createArchiveOrder(attrGoogleOrderNumber)
    createArchiveOrder = changeOrderState(attrGoogleOrderNumber, _
        "archive", "", "", "")
End Function


'*******************************************************************************
' The createChargeOrder function is a wrapper function that calls the
' changeOrderState function. The changeOrderState function, in turn,
' creates a <charge-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). 
' The <charge-order> command prompts Google Checkout to charge the customer 
' for an order and to change the order's financial order state to "CHARGING".
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemAmount               The amount that Google Checkout should 
'                                    charge the customer
' Returns:  <charge-order> XML
'*******************************************************************************
Function createChargeOrder(attrGoogleOrderNumber, elemAmount)
    createChargeOrder = changeOrderState(attrGoogleOrderNumber, _
        "charge", "", elemAmount, elemComment)
End Function


'*******************************************************************************
' The CreateCancelOrder function is a wrapper function that calls the
' ChangeOrderState function. The ChangeOrderState function, in turn,
' creates a <cancel-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). 
' The <cancel-order> command instructs Google Checkout to cancel an order.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemAmount               The reason an order is being canceled
' Input:    elemComment              A comment pertaining to a canceled order
' Returns:  <cancel-order> XML
'*******************************************************************************
Function createCancelOrder(attrGoogleOrderNumber, elemReason, elemComment)
    createCancelOrder = changeOrderState(attrGoogleOrderNumber, _
        "cancel", elemReason, "", elemComment)
End Function


'*******************************************************************************
' The createProcessOrder function is a wrapper function that calls the
' changeOrderState function. The changeOrderState function, in turn,
' creates a <process-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). 
' The <process-order> command changes the order's fulfillment order state
' to "PROCESSING".
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Returns:  <process-order> XML
'*******************************************************************************
Function createProcessOrder(attrGoogleOrderNumber)
    createProcessOrder = changeOrderState(attrGoogleOrderNumber, _
        "process", "", "", "")
End Function


'*******************************************************************************
' The createRefundOrder function is a wrapper function that calls the
' changeOrderState function. The changeOrderState function, in turn,
' creates a <refund-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber).
' The <refund-order> command instructs Google Checkout to issue a refund 
' for an order.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemReason               The reason an order is being refunded
' Input:    elemAmount               The amount that Google Checkout should
'                                    refund to the customer.
' Input:    elemComment              A comment pertaining to a refunded order
' Returns:  <refund-order> XML
'*******************************************************************************
Function createRefundOrder(attrGoogleOrderNumber, elemReason, elemAmount, _
    elemComment)

    createRefundOrder = changeOrderState(attrGoogleOrderNumber, _
        "refund", elemReason, elemAmount, elemComment)

End Function


'*******************************************************************************
' The createUnarchiveOrder function is a wrapper function that calls the
' changeOrderState function. The changeOrderState function, in turn,
' creates an <unarchive-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). The
' <unarchive-order> command moves an order from the Archive folder in the
' Google Checkout Merchant Center to the Inbox.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Returns:  <unarchive-order> XML
'*******************************************************************************
Function createUnarchiveOrder(attrGoogleOrderNumber)
    createUnarchiveOrder = changeOrderState(attrGoogleOrderNumber, _
        "unarchive", "", "", "")
End Function


'*******************************************************************************
' The createDeliverOrder function is a wrapper function that calls the
' changeShippingInfo function. The changeShippingInfo function, in turn,
' creates an <deliver-order> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber).
' The <deliver-order> command changes the order's fulfillment order state
' to "DELIVERED". It can also be used to add shipment tracking information
' for an order.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemCarrier              The carrier handling an order shipment
' Input:    elemTrackingNumber       The tracking number assigned to an
'                                    order shipment by the shipping carrier
' Returns:  <deliver-order> XML
'*******************************************************************************
Function createDeliverOrder(attrGoogleOrderNumber, elemCarrier, _
    elemTrackingNumber)

    createDeliverOrder = ChangeShippingInfo(attrGoogleOrderNumber, _
        "deliver-order", elemCarrier, elemTrackingNumber)

End Function


'*******************************************************************************
' The createAddTrackingData function is a wrapper function that calls the
' changeShippingInfo function. The changeShippingInfo function, in turn,
' creates an <add-tracking-data> command for the specified order, which is
' identified by its Google Checkout order number (attrGoogleOrderNumber). The
' <add-tracking-data> command adds shipment tracking information to an order.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemCarrier              The carrier handling an order shipment
' Input:    elemTrackingNumber       The tracking number assigned to an
'                                    order shipment by the shipping carrier
' Returns:  <add-tracking-data> XML
'*******************************************************************************
Function createAddTrackingData(attrGoogleOrderNumber, elemCarrier, _
    elemTrackingNumber)

    createAddTrackingData = ChangeShippingInfo(attrGoogleOrderNumber, _
        "add-tracking-data", elemCarrier, elemTrackingNumber)

End Function


'*******************************************************************************
' The changeOrderState function creates XML documents used to send
' Order Processing API commands to Google Checkout. This function 
' creates the XML for the following commands:
'           <archive-order> - moves order from Inbox to Archive folder
'           <cancel-order>  - cancels an order
'           <charge-order>  - charges an order and updated financial-order-state
'                             to "CHARGING"
'           <process-order> - changes fulfillment-order-state to "PROCESSING"
'           <refund-order>  - requests a refund for an order
'           <unarchive-order> - moves order from Archive folder to Inbox
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    functionName             The type of command that should be
'                                    created. Valid values for this
'                                    parameter are "archive", "cancel",
'                                    "charge", "process", "refund" and
'                                    "unarchive".
' Input:    elemReason               The reason an order is being refunded
' Input:    elemAmount               The amount that Google Checkout should
'                                    charge or refund to the customer.
' Input:    elemComment              A comment pertaining to an order
' Returns:  XML corresponding to the specified functionName
'*******************************************************************************
Function changeOrderState(attrGoogleOrderNumber, functionName, elemReason, _
    elemAmount, elemComment)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "changeOrderState(" & functionName & ")"

    ' Verify that the necessary parameter values have been provided.
    ' The attrGoogleOrderNumber and functionName parameters are
    ' required for all commands. The elemReason parameter is required
    ' for <cancel-order> and <refund-order> commands. In addition,
    ' if an elemAmount is provided for either the <charge-order> or
    ' <refund-order> commands, then the attrCurrency variable, which
    ' is defined in GlobalAPIFunctions.asp, must also have a value.
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrGoogleOrderNumber", _
        attrGoogleOrderNumber

    If functionName = "cancel" Or functionName = "refund" Then
        checkForError errorType, strFunctionName, "elemReason", elemReason
    End If
    
    ' Check for missing currency when amount is set
    If functionName = "charge" Or functionName = "refund" Then
        errorType = "MISSING_CURRENCY"
        If (elemAmount <> "") And (attrCurrency = "") Then
            errorHandler errorType, strFunctionName, "elemAmount", elemAmount
        End If
    End If


    ' Define the objects used to create the Order Processing API request
    Dim domOrderObj
    Dim domOrder
    Dim domReason
    Dim domAmount
    Dim domComment

    Set domOrderObj = Server.CreateObject(strMsxmlDomDocument)
    domOrderObj.async = False
    domOrderObj.appendChild(_
        domOrderObj.createProcessingInstruction("xml", strXmlVersionEncoding))

    ' Create the root tag for the Order Processing API command.
    ' Also set the "xmlns" and "google-order-number" attributes
    ' on that element.
    Set domOrder = domOrderObj.appendChild(_
        domOrderObj.createElement(functionName & "-order"))
    domOrder.setAttribute "xmlns", strXmlns
    domOrder.setAttribute "google-order-number", attrGoogleOrderNumber

    ' Add <reason> element to <cancel-order> and <refund-order> commands
    If functionName = "cancel" Or functionName = "refund" Then
        Set domReason = _
            domOrder.appendChild(domOrderObj.createElement("reason"))
        domReason.Text = elemReason
    End If

    ' Add <amount> element to <charge-order> and <refund-order> commands
    If functionName = "charge" Or functionName = "refund" Then
        If elemAmount <> "" Then
            Set domAmount = _
                domOrder.appendChild(domOrderObj.createElement("amount"))
            domAmount.setAttribute "currency", attrCurrency
            domAmount.Text = elemAmount
        End If
    End If

    ' Add <comment> element
    If elemComment <> "" Then
        Set domComment = _
            domOrder.appendChild(domOrderObj.createElement("comment"))
        domComment.Text = elemComment
    End If

    changeOrderState = domOrderObj.xml

    ' Release the objects used to create the Order Processing API request
    Set domOrderObj = Nothing
    Set domOrder = Nothing
    Set domReason = Nothing
    Set domAmount = Nothing
    Set domComment = Nothing

End Function


'*******************************************************************************
' The changeShippingInfo function creates XML documents used to send
' Order Processing API commands to Google Checkout. This function creates 
' the XML for the following commands:
'         <deliver-order>
'         <add-tracking-data>
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    functionName             The type of command that should be
'                                    created. Valid values for this
'                                    parameter are "deliver" and 
'                                    "add-tracking-data".
' Input:    elemCarrier              The carrier handling an order shipment.
' Input:    elemTrackingNumber       The tracking number assigned to an
'                                    order shipment by the shipping carrier
' Returns:  XML corresponding to the specified functionName
'*******************************************************************************
Function changeShippingInfo(attrGoogleOrderNumber, functionName, _
    elemCarrier, elemTrackingNumber)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "changeShippingInfo(" & functionName & ")"

    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrGoogleOrderNumber", _
        attrGoogleOrderNumber
    
    ' Check for missing tracking number when carrier is set
    ' Verify that the necessary parameter values have been provided.
    ' The attrGoogleOrderNumber and functionName parameters are
    ' required for all commands. For the <deliver-order> command, the
    ' elemCarrier and elemTrackingNumber parameters are optional; however,
    ' if the elemCarrier is provided, then a elemTrackingNumber must also
    ' be provided. For the <add-tracking-data> command, the elemCarrier
    ' and elemTrackingNumber parameters are both required.
    If functionName = "deliver-order" Then
        errorType = "MISSING_TRACKING"
        If (elemCarrier <> "") And (elemTrackingNumber) = "" Then
            errorHandler errorType, strFunctionName, "elemCarrier", elemCarrier
        End If
    ElseIf functionName = "add-tracking-data" Then
        checkForError errorType, strFunctionName, "elemCarrier", elemCarrier
        checkForError errorType, strFunctionName, "elemTrackingNumber", _
            elemTrackingNumber
    End If

    ' Define the objects used to create the Order Processing API request
    Dim domObj
    Dim domShipping
    Dim domTrackingData
    Dim domCarrier
    Dim domComment

    Set domObj = Server.CreateObject(strMsxmlDomDocument)
    domObj.async = False
    domObj.appendChild(domObj.createProcessingInstruction("xml", _
        strXmlVersionEncoding))

    ' Create the root tag for the Order Processing API command.
    ' Also set the "xmlns" and "google-order-number" attributes
    ' on that element.
    Set domShipping = domObj.appendChild(domObj.createElement(functionName))
    domShipping.setAttribute "xmlns", strXmlns
    domShipping.setAttribute "google-order-number", attrGoogleOrderNumber

    ' Add the <carrier> and <tracking-number> elements
    If elemCarrier <> "" Then

        Set domTrackingData = _
            domShipping.appendChild(domObj.createElement("tracking-data"))

        Set domCarrier = _
            domTrackingData.appendChild(domObj.createElement("carrier"))
        domCarrier.Text = elemCarrier

        Set domComment = _
            domTrackingData.appendChild(domObj.createElement("tracking-number"))
        domComment.Text = elemTrackingNumber

    End If

    changeShippingInfo = domObj.xml

    ' Release the objects used to create the Order Processing API request
    Set domObj = Nothing
    Set domShipping = Nothing
    Set domTrackingData = Nothing
    Set domCarrier = Nothing
    Set domComment = Nothing

End Function


'*******************************************************************************
' The createAddMerchantOrderNumber function creates the XML for the
' <add-merchant-order-number> Order Processing API command. This command
' is used to associate the Google order number with the ID that the
' merchant assigns to the same order.
'
' Inputs:   attrGoogleOrderNumber    A number, assigned by Google Checkout, that
'                                    that uniquely identifies an order.
'           elemMerchantOrderNumber  A string, assigned by the merchant, 
'                                    that uniquely identifies the order.
' Returns:  <add-merchant-order-number> XML
'*******************************************************************************
Function createAddMerchantOrderNumber(attrGoogleOrderNumber, _
    elemMerchantOrderNumber)

    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAddMerchantOrderNumber()"

    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrGoogleOrderNumber", _
        attrGoogleOrderNumber
    checkForError errorType, strFunctionName, "elemMerchantOrderNumber", _
        elemMerchantOrderNumber

    ' Define the objects used to create the <add-merchant-order-number> command
    Dim domObj
    Dim domAddMerchantOrderNumber
    Dim domMerchantOrderNumber

    Set domObj = Server.CreateObject(strMsxmlDomDocument)
    domObj.async = False
    domObj.appendChild(domObj.createProcessingInstruction("xml", _
        strXmlVersionEncoding))

    ' Create the root tag for the Order Processing API command.
    ' Also set the "xmlns" and "google-order-number" attributes
    ' on that element.
    Set domAddMerchantOrderNumber = domObj.appendChild( _
        domObj.createElement("add-merchant-order-number"))

    domAddMerchantOrderNumber.setAttribute "xmlns", strXmlns

    domAddMerchantOrderNumber.setAttribute "google-order-number", _
        attrGoogleOrderNumber

    ' Add the <merchant-order-number> element
    Set domMerchantOrderNumber = domAddMerchantOrderNumber.appendChild( _
        domObj.createElement("merchant-order-number"))

    domMerchantOrderNumber.Text = elemMerchantOrderNumber

    createAddMerchantOrderNumber = domObj.xml
End Function


'*******************************************************************************
' The CreateSendBuyerMessage function creates the XML for the
' <send-buyer-message> Order Processing API command. This command
' is used to send a message to a customer.
'
' Input:    attrGoogleOrderNumber    A number, assigned by Google Checkout,
'                                    that uniquely identifies an order.
' Input:    elemMessage              The text of the message that you
'                                    want to send to the customer
' Input:    elemSendEmail            A Boolean value that indicates whether
'                                    Google Checkout should email the customer 
'                                    when the <deliver-order> command is
'                                    processed for the order.
' Returns:  <send-buyer-message> XMLDOM
'*******************************************************************************
Function createSendBuyerMessage(attrGoogleOrderNumber, elemMessage, _
    elemSendEmail)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createSendBuyerMessage()"

    ' The attrGoogleOrderNumber and elemMessage parameters must both have values
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrGoogleOrderNumber", _
        attrGoogleOrderNumber
    checkForError errorType, strFunctionName, "elemMessage", elemMessage

    ' Define the objects used to create the <send-buyer-message> command
    Dim domObj
    Dim domSendBuyerMessage
    Dim domMessage
    Dim domSendEmail


    ' Create the root element for the <send-buyer-message> command
    ' Also set the "xmlns" and "google-order-number" attributes
    ' on that element.
    Set domObj = Server.CreateObject(strMsxmlDomDocument)
    domObj.async = False
    domObj.appendChild(domObj.createProcessingInstruction("xml", _
        strXmlVersionEncoding))

    Set domSendBuyerMessage = _
        domObj.appendChild(domObj.createElement("send-buyer-message"))
    domSendBuyerMessage.setAttribute "xmlns", strXmlns
    domSendBuyerMessage.setAttribute "google-order-number", _
        attrGoogleOrderNumber

    ' Add the <message> element to the command
    Set domMessage = _
        domSendBuyerMessage.appendChild(domObj.createElement("message"))
    domMessage.Text = elemMessage

    ' Add the <send-email> element to the command
    If elemSendEmail <> "" Then
        Set domSendEmail = _
            domSendBuyerMessage.appendChild(domObj.createElement("send-email"))
        domSendEmail.Text = elemSendEmail
    End If

    createSendBuyerMessage = domObj.xml

    ' Release the objects used to create the <send-buyer-message> command
    Set domObj = Nothing
    Set domSendBuyerMessage = Nothing
    Set domMessage = Nothing
    Set domSendEmail = Nothing

End Function

%>