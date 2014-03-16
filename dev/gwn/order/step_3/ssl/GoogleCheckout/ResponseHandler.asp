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
' This sample code is a simple listener for the responses that the 
' Google Checkout server sends to the merchants. When a response is 
' received, it logs the message to a local file and processes the message.
'*******************************************************************************

' Include ASP libraries used to process Google Checkout responses
%>

<!--#INCLUDE file="GlobalAPIFunctions.asp"--> 
<!--#INCLUDE file="ResponseHandlerAPIFunctions.asp"--> 
<!--#INCLUDE file="MerchantCalculationsAPIFunctions.asp"--> 

<%

' Define objects used to process Google Checkout response
Dim xmlResponse
Dim xmlAcknowledgment

' Retrieve the XML sent in the HTTP POST request to the ResponseHandler
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
Dim nIndex
For nIndex = 1 to LenB(biData)
    xmlResponse = xmlResponse & Chr(AscB(MidB(biData,nIndex,1)))
Next

' Log the XML received in the HTTP request
logMessage logFilename, xmlResponse

' Call the processXmlData function, which is defined in
' ResponseHandlerAPIFunctions.asp. The processXmlData will route
' the XML data to the function that handles the particular type
' of XML message contained in the POST request.
processXmlData(xmlResponse)

%>