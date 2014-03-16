<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.5

'@FILENAME: incverify.asp
	 



'@DESCRIPTION: Verify the order information

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde  Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'------------------------------------------------------------------
' Checks PhoneFax type
' Returns 0 for recorded, 1 for non-recorded
'------------------------------------------------------------------

Function CheckPaymentMethod(sPayMethod)

Dim sLocalSQL, rsTransType, sReturn, sTempTransType

	sLocalSQL = "SELECT DISTINCT transtype FROM sfTransactionTypes WHERE transType='" & makeInputSafe(sPayMethod) & "' and transIsActive = 1"
	
	Set rsTransType = CreateObject("ADODB.RecordSet")
	rsTransType.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
	
	If rsTransType.EOF or rsTransType.BOF Then
		sReturn = 0
	Else	
		sReturn = 1
	End If
	
	closeObj(rsTransType)
	CheckPayMentMethod = sReturn
	
End Function

%>
