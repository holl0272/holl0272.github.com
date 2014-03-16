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
' "NotificationAPIFunctions.asp" is a client library of functions 
' that enable merchants to systematically handle Google Checkout notifications
' using the Notification API. You will need to modify these functions
' to take the appropriate actions when you receive notifications.
'*******************************************************************************

'******** Functions for processing asynchronous notification messages *********

'*******************************************************************************
' The processNewOrderNotification function is a shell function for
' handling a <new-order-notification>. You will need to modify this
' function to transfer the information contained in a
' <new-order-notification> to your internal systems that process that data.
'
' Input:   domResponseObj    <new-order-notification> XML DOM
'*******************************************************************************
Function processNewOrderNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' New order notifications inform you of new orders that have
    ' been submitted through Google Checkout. A <new-order-notification>
    ' message contains a list of the items in an order, the tax
    ' assessed on the order, the shipping method selected for the
    ' order and the shipping address for the order.
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <new-order-notification> to your internal systems that
    ' process this order data.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The processOrderStateChangeNotification function is a shell function
' for handling an <order-state-change-notification>. You will need to
' modify this function to transfer the information contained in a
' <order-state-change-notification> to your internal systems that
' process that data.
'
' Input:  domResponseObj    <order-state-change-notification> XML DOM
'*******************************************************************************
Function processOrderStateChangeNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' Order state change notifications signal an update to an order's
    ' financial status or its fulfillment status. An
    ' <order-state-change-notification> identifies the new financial
    ' and fulfillment statuses for an order. It also identifies the
    ' previous statuses for the order. Google Checkout will send an
    ' <order-state-change-notification> to confirm status changes that
    ' you trigger by using the Order Processing API requests. For
    ' example, if you send Google Checkout a <cancel-order> request, 
    ' it will respond through the Notification API to inform you that the
    ' order's status has been changed to "canceled".
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <order-state-change-notification> to your internal systems that
    ' process financial or fulfillment status information.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The processChargeAmountNotification function is a shell function for
' handling a <charge-amount-notification>. You will need to modify this
' function to relay the information in the <charge-amount-notification>
' to your internal systems that process that data.
'
' Input:  domResponseObj    <charge-amount-notification> XML DOM
'*******************************************************************************
Function processChargeAmountNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' Charge amount notifications inform you that a customer has been
    ' charged for either the full amount or a partial amount of an
    ' order. A <charge-amount-notification> contains the order number
    ' that Google assigned to the order, the value of the most recent
    ' charge to the customer and the total amount that has been
    ' charged to the customer for the order. Google Checkout will send a
    ' <charge-amount-notification> after charging the customer.
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <charge-amount-notification> to your internal systems that
    ' process this order data.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The processChargebackAmountNotification function is a shell function
' for handling a <chargeback-amount-notification>. You will need to
' modify this function to transfer the information contained in a
' <chargeback-amount-notification> to your internal systems that
' process that data.
'
' Input:  domResponseObj    <chargeback-amount-notification> XML DOM
'*******************************************************************************
Function processChargebackAmountNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' Chargeback amount notifications inform you that a customer
    ' has initiated a chargeback against an order and that Google Checkout has
    ' approved the chargeback. A <chargeback-amount-notification>
    ' contains the order number that Google assigned to the order,
    ' the value of the most recent chargeback against the order
    ' and the total amount that has been charged back against the
    ' order. Google Checkout will send a <chargeback-amount-notification>
    ' after approving the chargeback.
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <chargeback-amount-notification> to your internal systems that
    ' process this order data.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The processRefundAmountNotification function is a shell function for
' handling a <refund-amount-notification>. You will need to modify this
' function to transfer the information contained in a
' <refund-amount-notification> to your internal systems that handle that data.
'
' Input:  domResponseObj    <refund-amount-notification> XML DOM
'*******************************************************************************
Function processRefundAmountNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' Refund amount notifications inform you that a customer has been
    ' refunded either the full amount or a partial amount of an order
    ' total. A <refund-amount-notification> contains the order number
    ' that Google assigned to the order, the value of the most recent
    ' refund to the customer and the total amount that has been
    ' refunded to the customer for the order. Google Checkout will send a
    ' <refund-amount-notification> after refunding the customer.
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <refund-amount-notification> to your internal systems that
    ' process this order data.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The processRiskInformationNotification function is a shell function for
' handling a <risk-information-notification>. You will need to modify this
' function to transfer the information contained in a
' <risk-information-notification> to your internal systems that process
' that data.
'
' Input:  domResponseObj    <risk-information-notification> XML DOM
'*******************************************************************************
Function processRiskInformationNotification(domResponseObj)
    ' +++ CHANGE ME +++
    ' Risk information notifications provide financial information about
    ' a transaction to help you ensure that an order is not fraudulent.
    ' A <risk-information-notification> includes the customer's billing
    ' address, a partial credit card number and other values to help you
    ' verify that an order is not fraudulent. Google Checkout will send you a
    ' <risk-information-notification> message after completing its
    ' risk analysis on a new order.
    '
    ' If you are implementing the Notification API, you need to
    ' modify this function to relay the information in the
    ' <risk-information-notification> to your internal systems that
    ' process this order data.
    sendNotificationAcknowledgment
End Function


'*******************************************************************************
' The sendNotificationAcknowledgment function responds to a Google Checkout
' notification with a <notification-acknowledgment> message. If you do
' not send a <notification-acknowledgment> in response to a Google Checkout
' notification, Google Checkout will resend the notification multiple times.
'*******************************************************************************
Function sendNotificationAcknowledgment

    ' Respond to the notification with <notification-acknowledgment>
    xmlAcknowledgment = _
      "<?xml version=""1.0"" encoding=""UTF-8""?>" _
      & "<notification-acknowledgment " _
      & "xmlns=""" & strXmlns & """/>"
    response.write xmlAcknowledgment
    
    ' Log <notification-acknowledgment>
    logMessage logFilename, xmlAcknowledgment

End Function
%>
