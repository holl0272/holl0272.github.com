<%
'********************************************************************************
'*   Sandshot Software Product Review Component									*
'*   Release Version   1.00.001													*
'*   Release Date      March 15, 2005											*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim cstrSendNotificationEmailTo
Dim contentAuthorEmail
Dim contentAuthorName
Dim contentAuthorRating
Dim contentAuthorShowEmail
Dim contentContent
Dim mblnRequiresProcessing
Dim mlngNumReviews
Dim mlngAvgReviewScore
Dim mstrErrorMessage_ProductReview

	mblnRequiresProcessing = True
	cstrSendNotificationEmailTo = adminPrimaryEmail	'leave blank to not send

'**********************************************************
'*	Functions
'**********************************************************

'Sub castFoundUsefulVote(byVal lngcontentID, byVal lngRating)
'Function getFoundUsefulVotes(byVal lngcontentID, byRef aryVotes)
'Function loadProductReviews(byVal strProductID, byRef aryProductReviews)
'Sub loadProductReviewSummary(byVal strProductID)
'Function ratingDisplayImage(byVal dblRating)
'Function saveReview(byVal strProductID)
'Sub sendProductReviewNotificationEmail(byVal strProductID)
'Sub ssProductReview(byVal strProductID)
'Function submitProductReview(byVal strProductID)
'Sub writeProductReviewDetails(byVal strProductID)
'Sub writeProductReviewSummary(byVal strProductID)
'Function validateReview(byVal strProductID)
'Sub WriteReviewForm(byVal strProductID)

'**********************************************************
'*	Begin Page Code
'**********************************************************

Sub castFoundUsefulVote(byVal lngcontentID, byVal lngRating)

Dim pobjCmd

	If Len(lngcontentID) = 0 Or Not isNumeric(lngcontentID) Then Exit Sub
	If Len(lngRating) = 0 Or Not isNumeric(lngRating) Then Exit Sub

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Insert Into contentFoundUseful (contentFoundUsefulContentID, contentFoundUsefulScore) Values (?, ?)"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("contentFoundUsefulContentID", adInteger, adParamInput, 4, lngcontentID)
		.Parameters.Append .CreateParameter("contentFoundUsefulScore", adInteger, adParamInput, 4, lngRating)
		.Execute,, 128
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Sub	'castFoundUsefulVote

'***************************************************************************************************************************************************************************************

Function getFoundUsefulVotes(byVal lngcontentID, byRef aryVotes)

Dim pblnSuccess
Dim pobjCmd
Dim pobjRS

	pblnSuccess = False
	
	If Len(lngcontentID) = 0 Or Not isNumeric(lngcontentID) Then
		getFoundUsefulVotes = pblnSuccess
		Exit Function
	End If

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select contentFoundUsefulScore From contentFoundUseful Where contentFoundUsefulContentID=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("contentFoundUsefulContentID", adInteger, adParamInput, 4, lngcontentID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			aryVotes = pobjRS.GetRows()
			pblnSuccess = True
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	getFoundUsefulVotes = pblnSuccess

End Function	'getFoundUsefulVotes

'***************************************************************************************************************************************************************************************

Function loadProductReviews(byVal strProductID, byRef aryProductReviews)

Dim pblnResult
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim paryVotes
Dim i,j

	mlngNumReviews = 0
	mlngAvgReviewScore = 0
	pblnResult = False

	If Len(strProductID) = 0 Or Len(strProductID) > 50 Then
		loadProductReviews = pblnResult
		Exit Function
	End If

	pstrSQL = "SELECT contentID, contentContent, contentAuthorName, contentAuthorEmail, contentAuthorShowEmail, contentAuthorRating, contentDateCreated, 0 as foundUseful, 0 as foundWorthless" _
			& " FROM sfProducts RIGHT JOIN content ON sfProducts.sfProductID = content.contentReferenceID" _
			& " WHERE ((content.contentApprovedForDisplay=1 Or content.contentApprovedForDisplay=-1) AND (content.contentContentType=5) AND (sfProducts.prodID=?))" _
			& " ORDER BY contentDateCreated DESC"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProductID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			aryProductReviews = pobjRS.GetRows()

			mlngNumReviews = UBound(aryProductReviews, 2) + 1
			For i = 0 To mlngNumReviews - 1
				mlngAvgReviewScore = mlngAvgReviewScore + aryProductReviews(5, i)
				If getFoundUsefulVotes(aryProductReviews(0, i), paryVotes) Then
					For j = 0 To UBound(paryVotes, 2)
						If paryVotes(0, j) = 0 Then
							aryProductReviews(8, i) = aryProductReviews(8, i) + 1
						Else
							aryProductReviews(7, i) = aryProductReviews(7, i) + 1
						End If
					Next 'j
				End If
			Next 'i
			If mlngNumReviews <> 0 Then mlngAvgReviewScore = mlngAvgReviewScore / mlngNumReviews

			pblnResult = True
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	loadProductReviews = pblnResult
	
End Function	'loadProductReviews

'***************************************************************************************************************************************************************************************

Function loadProductReviewsByCustomer(byVal lngCustomerID, byRef aryProductReviews)

Dim pblnResult
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim paryVotes
Dim i,j

	mlngNumReviews = 0
	mlngAvgReviewScore = 0
	pblnResult = False

	If Len(lngCustomerID) = 0 Or Not isNumeric(lngCustomerID) Then
		loadProductReviewsByCustomer = pblnResult
		Exit Function
	End If

	pstrSQL = "SELECT contentID, contentContent, contentAuthorName, contentAuthorEmail, contentAuthorShowEmail, contentAuthorRating, contentDateCreated, 0 as foundUseful, 0 as foundWorthless, prodID, prodName" _
			& " FROM sfProducts RIGHT JOIN content ON sfProducts.sfProductID = content.contentReferenceID" _
			& " WHERE ((content.contentApprovedForDisplay=1 Or content.contentApprovedForDisplay=-1) AND (content.contentContentType=5) AND (content.contentAuthorID=?))" _
			& " ORDER BY contentDateCreated DESC"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adInteger, adParamInput, 4, lngCustomerID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			aryProductReviews = pobjRS.GetRows()

			mlngNumReviews = UBound(aryProductReviews, 2) + 1
			For i = 0 To mlngNumReviews - 1
				mlngAvgReviewScore = mlngAvgReviewScore + aryProductReviews(5, i)
				If getFoundUsefulVotes(aryProductReviews(0, i), paryVotes) Then
					For j = 0 To UBound(paryVotes, 2)
						If paryVotes(0, j) = 0 Then
							aryProductReviews(8, i) = aryProductReviews(8, i) + 1
						Else
							aryProductReviews(7, i) = aryProductReviews(7, i) + 1
						End If
					Next 'j
				End If
			Next 'i
			If mlngNumReviews <> 0 Then mlngAvgReviewScore = mlngAvgReviewScore / mlngNumReviews

			pblnResult = True
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	loadProductReviewsByCustomer = pblnResult
	
End Function	'loadProductReviewsByCustomer

'***************************************************************************************************************************************************************************************

Sub loadProductReviewSummary(byVal strProductID)

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	mlngNumReviews = 0
	mlngAvgReviewScore = 0

	If Len(strProductID) = 0 Or Len(strProductID) > 50 Then Exit Sub

	pstrSQL = "SELECT Count(content.contentID) AS CountOfcontentID, Sum(content.contentAuthorRating) AS SumOfcontentAuthorRating" _
			& " FROM sfProducts INNER JOIN content ON sfProducts.sfProductID = content.contentReferenceID" _
			& " GROUP BY content.contentApprovedForDisplay, content.contentContentType, sfProducts.prodID" _
			& " HAVING ((content.contentApprovedForDisplay=1 Or content.contentApprovedForDisplay=-1) AND (content.contentContentType=5) AND (sfProducts.prodID=?))"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProductID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			mlngNumReviews = pobjRS.Fields("CountOfcontentID").Value
			If mlngNumReviews > 0 Then mlngAvgReviewScore = pobjRS.Fields("SumOfcontentAuthorRating").Value / mlngNumReviews
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
End Sub	'loadProductReviewSummary

'***************************************************************************************************************************************************************************************

Function ratingDisplayImage(byVal dblRating)

	If dblRating >= 4.5 Then
		ratingDisplayImage = "<img src=""images/stars5.jpg"" alt=""5 stars"">"
	ElseIf dblRating >= 3.5 Then
		ratingDisplayImage = "<img src=""images/stars4.jpg"" alt=""4 stars"">"
	ElseIf dblRating >= 2.5 Then
		ratingDisplayImage = "<img src=""images/stars3.jpg"" alt=""3 stars"">"
	ElseIf dblRating >= 1.5 Then
		ratingDisplayImage = "<img src=""images/stars2.jpg"" alt=""2 stars"">"
	ElseIf dblRating >= 0.5 Then
		ratingDisplayImage = "<img src=""images/stars1.jpg"" alt=""1 stars"">"
	Else
		ratingDisplayImage = "<img src=""images/stars0.jpg"" alt=""0 stars"">"
	End If
	
End Function	'ratingDisplayImage

'***************************************************************************************************************************************************************************************

Function saveReview(byVal strProductID)

Dim pobjCmd
Dim pobjRS
Dim pstrResult
Dim sfProductID

	'Need to retrieve productID from sfProducts table

	If Len(strProductID) = 0 Or Len(strProductID) > 50 Then
		pstrResult = pstrResult & "Invalid product to review"
		saveReview = pstrResult
		Exit Function
	End If

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "SELECT sfProductID FROM sfProducts WHERE (sfProducts.prodID=?)"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProductID)
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			pstrResult = pstrResult & "Invalid product to review"
		Else
			sfProductID = pobjRS.Fields("sfProductID").Value
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
	End With	'pobjCmd
	Set pobjCmd = Nothing

	'Check to make sure a review doesn't already exist by this customer
	If Len(pstrResult) = 0 And visitorCustomerID <> 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "SELECT contentID, contentApprovedForDisplay, contentDateCreated FROM content WHERE contentAuthorID=? AND contentContentType=? AND contentReferenceID=?"
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("contentAuthorID", adInteger, adParamInput, 4, visitorCustomerID)
			.Parameters.Append .CreateParameter("contentContentType", adInteger, adParamInput, 4, 5)
			.Parameters.Append .CreateParameter("contentReferenceID", adInteger, adParamInput, 4, sfProductID)
			Set pobjRS = .Execute
			
			If Not pobjRS.EOF Then
				If Abs(pobjRS.Fields("contentApprovedForDisplay").Value) = 1 Then
					pstrResult = pstrResult & "You already reviewed this product on " & pobjRS.Fields("contentDateCreated").Value & "."
				Else
					pstrResult = pstrResult & "Your prior review of this product on " & pobjRS.Fields("contentDateCreated").Value & " is being processed."
				End If
			End If	'pobjRS.EOF
			pobjRS.Close
			Set pobjRS = Nothing
			
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(pstrResult) = 0
	
	If Len(pstrResult) = 0 Then
	
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Insert Into content (contentAuthorID, contentContentType, contentReferenceID, contentApprovedForDisplay, contentContent, contentAuthorName, contentAuthorEmail, contentAuthorShowEmail, contentAuthorRating, contentDateCreated, contentDateModified) Values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
			Set .ActiveConnection = cnn
			
			'this check added because of a report of an error; unable to duplicate and it shouldn't happen but . . .
			If Len(visitorCustomerID) = 0 Or Not isNumeric(visitorCustomerID) Then
				.Parameters.Append .CreateParameter("contentAuthorID", adInteger, adParamInput, 4, 0)
			Else
				.Parameters.Append .CreateParameter("contentAuthorID", adInteger, adParamInput, 4, visitorCustomerID)
			End If
			
			.Parameters.Append .CreateParameter("contentContentType", adInteger, adParamInput, 4, 5)
			.Parameters.Append .CreateParameter("contentReferenceID", adInteger, adParamInput, 4, sfProductID)
			
			If mblnRequiresProcessing Then
				.Parameters.Append .CreateParameter("contentApprovedForDisplay", adInteger, adParamInput, 4, 0)
			Else
				.Parameters.Append .CreateParameter("contentApprovedForDisplay", adInteger, adParamInput, 4, 1)
			End If

			If Len(contentContent) = 0 Then
				.Parameters.Append .CreateParameter("contentContent", adLongVarWChar, adParamInput, 2147483646, NULL)
			Else
				.Parameters.Append .CreateParameter("contentContent", adLongVarWChar, adParamInput, 2147483646, contentContent)
			End If
			
			If Len(contentAuthorName) > 100 Then
				.Parameters.Append .CreateParameter("contentAuthorName", adVarChar, adParamInput, 100, Left(contentAuthorName, 100))
			Else
				.Parameters.Append .CreateParameter("contentAuthorName", adVarChar, adParamInput, 100, contentAuthorName)
			End If
			
			If Len(contentAuthorEmail) > 100 Then
				.Parameters.Append .CreateParameter("contentAuthorEmail", adVarChar, adParamInput, 100, Left(contentAuthorEmail, 100))
			Else
				.Parameters.Append .CreateParameter("contentAuthorEmail", adVarChar, adParamInput, 100, contentAuthorEmail)
			End If
			
			.Parameters.Append .CreateParameter("contentAuthorShowEmail", adInteger, adParamInput, 4, contentAuthorShowEmail)
			.Parameters.Append .CreateParameter("contentAuthorRating", adInteger, adParamInput, 4, contentAuthorRating)
			.Parameters.Append .CreateParameter("contentDateCreated", adDBTimeStamp, adParamInput, 16, Now())
			.Parameters.Append .CreateParameter("contentDateModified", adDBTimeStamp, adParamInput, 16, Now())
			
			On Error Resume Next
			.Execute,,128
			
			If Err.number <> 0 Then
				pstrResult = "Error saving review. Please submit your review again."
				pstrResult = pstrResult & "<br />Error " & err.number & ": " & err.Description & "<br />"
			Else
				Call sendProductReviewNotificationEmail(strProductID)
			End If
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(pstrResult) = 0

	saveReview = pstrResult
		
End Function	'saveReview

'***************************************************************************************************************************************************************************************

Sub sendProductReviewNotificationEmail(byVal strProductID)

Dim pstrEmailBody
Dim pstrEmailSubject

	If Len(cstrSendNotificationEmailTo) = 0 Then Exit Sub

	If mblnRequiresProcessing Then
		pstrEmailSubject = "Product review, approval required, product " & strProductID
	Else
		pstrEmailSubject = "Product review, product " & strProductID
	End If

	pstrEmailBody = "Product " & strProductID & " was reviewed." & vbcrlf _
				  & "Reviewer: " & contentAuthorName & vbcrlf _
				  & "          " & contentAuthorEmail & vbcrlf _
				  & "Rating: " & contentAuthorRating & vbcrlf _
				  & "        " & contentContent & vbcrlf
				  
	If mblnRequiresProcessing Then
		pstrEmailBody = pstrEmailBody & vbcrlf & "Approval is required for it to show up." & vbcrlf
	Else
		pstrEmailBody = pstrEmailBody & vbcrlf & "It has been automatically approved for viewing." & vbcrlf
	End If

	pstrEmailBody = pstrEmailBody & vbcrlf & adminDomainName & "ssl/ssAdmin/ssProductReviewsAdmin.asp" & VbCrLf

	Call createMail("-", cstrSendNotificationEmailTo & "|" & contentAuthorEmail & "|" & "-" & "|" & pstrEmailSubject & "|" & pstrEmailBody)

End Sub	'sendProductReviewNotificationEmail

'***************************************************************************************************************************************************************************************

Sub ssProductReview(byVal strProductID)

Dim pstrReviewAction
Dim plngContentReferenceID
pstrReviewAction = LoadRequestValue("ReviewAction")

'Options are:
'Nothing - view summary
'ViewDetails - view details
'WriteReview - initial entry
'SubmitReview - 
'RateReview

	If Err.number <> 0 Then Err.Clear

	Select Case pstrReviewAction
		Case "RateReviewYes"
			plngContentReferenceID = LoadRequestValue("contentReferenceID")
			Call castFoundUsefulVote(plngContentReferenceID, 1)
			mstrErrorMessage_ProductReview = "<font color=red><b>Thank you! We're pleased you found this review useful.</b></font>"
			Call writeProductReviewDetails(strProductID)
		Case "RateReviewNo"
			plngContentReferenceID = LoadRequestValue("contentReferenceID")
			Call castFoundUsefulVote(plngContentReferenceID, 0)
			mstrErrorMessage_ProductReview = "<font color=red><b>Thank you! We're sorry this review wan't useful.</b></font>"
			Call writeProductReviewDetails(strProductID)
		Case "ReportReview"
		
		Case "ReadReviews"
			Call writeProductReviewDetails(strProductID)
		Case "SubmitReview"
			If submitProductReview(strProductID) Then
				If mblnRequiresProcessing Then
					mstrErrorMessage_ProductReview = "<font color=red><b>Thank you! You're review has been submitted and will appear shortly.</b></font>"
				Else
					mstrErrorMessage_ProductReview = "<font color=red><b>Thank you! You're review has been submitted.</b></font>"
				End If
				Call writeProductReviewSummary(strProductID)
			Else
				Call WriteReviewForm(strProductID)
			End If
		Case "ViewDetails"
		
		Case "WriteReview"
			Call WriteReviewForm(strProductID)
		Case Else
			Call writeProductReviewSummary(strProductID)
	End Select

End Sub	'ssProductReview

'***************************************************************************************************************************************************************************************

Function submitProductReview(byVal strProductID)

Dim pstrSQL

	contentAuthorName = Trim(Request.Form("contentAuthorName"))
	contentAuthorEmail = Trim(Request.Form("contentAuthorEmail"))
	contentAuthorRating = Trim(Request.Form("contentAuthorRating"))
	contentContent = Trim(Request.Form("contentContent"))
	contentAuthorShowEmail = Trim(Request.Form("contentAuthorShowEmail"))
	If contentAuthorShowEmail <> "1" Then contentAuthorShowEmail = 0
	
	'Now validate the required fields
	mstrErrorMessage_ProductReview = validateReview(strProductID)
	
	If Len(mstrErrorMessage_ProductReview) = 0 Then
		mstrErrorMessage_ProductReview = saveReview(strProductID)
	Else
		mstrErrorMessage_ProductReview = "<font color=red>There was an error processing your review. Please correct the following items:<ul>" & mstrErrorMessage_ProductReview & "</ul>"
	End If
	
	submitProductReview = CBool(Len(mstrErrorMessage_ProductReview) = 0)

End Function	'submitProductReview

'***************************************************************************************************************************************************************************************

Sub writeProductReviewDetails(byVal strProductID)

Dim i
Dim paryProductReviews
	'0-contentReferenceID
	'1-contentContent
	'2-contentAuthorName
	'3-contentAuthorEmail
	'4-contentAuthorShowEmail
	'5-contentAuthorRating
	'6-contentDateCreated
Dim pstrAuthor
Dim pstrURL

	pstrURL = CurrentPage & "?ReviewAction=WriteReview&amp;PRODUCT_ID=" & strProductID
	If Not loadProductReviews(strProductID, paryProductReviews) Then
		mstrErrorMessage_ProductReview = "Error loading product reviews"
	End If
	%>
	<% If Len(mstrErrorMessage_ProductReview) > 0 Then Response.Write "<p>" & mstrErrorMessage_ProductReview & "</p>" %>
	<table class="Section" border="1" cellpadding="0" cellspacing="0">
	  <tr>
		<th class="tdMiddleTopBanner">Product Review</th>
	  </tr>
	  <tr>
		<td class="tdContent2">
		<%
		If mlngNumReviews = 0 Then
			Response.Write "This product has not been rated yet."
		Else
			If mlngNumReviews = 1 Then
				Response.Write "One reviewer gave this a rating of " & ratingDisplayImage(mlngAvgReviewScore)
			Else
				Response.Write mlngNumReviews & " reviewers gave this an average rating of " & ratingDisplayImage(mlngAvgReviewScore)
			End If

			For i = 0 To mlngNumReviews - 1
			
				If Abs(paryProductReviews(4, i)) = 1 Then
					pstrAuthor = "<a href='mailTo:" & Trim(paryProductReviews(3, i)) & "'>" & Trim(paryProductReviews(2, i)) & "</a>"
				Else
					pstrAuthor = Trim(paryProductReviews(2, i))
				End If
				
				Response.Write "<hr />"
				Call DisplayProductReviewEditLink(paryProductReviews, i)
				Response.Write pstrAuthor & " " & ratingDisplayImage(paryProductReviews(5, i)) & "<br />"
				If Len(Trim(paryProductReviews(1, i))) > 0 Then
					Response.Write "<strong>Comments:</strong> " & Trim(paryProductReviews(1, i)) & "<br />"
				End If
				Response.Write "<strong>Date Reviewed:</strong> " & FormatDateTime(paryProductReviews(6, i)) & "<br />"
				
				If paryProductReviews(7, i) > 0 Or paryProductReviews(8, i) > 0 Then
					Response.Write paryProductReviews(7, i) & " of " & paryProductReviews(7, i) + paryProductReviews(8, i) & " found this review useful.<br />"
				End If
				Response.Write "Did you find this review helpful?&nbsp;&nbsp;" _
							 & "<a href=""" & CurrentPage & "?ReviewAction=RateReviewYes&amp;PRODUCT_ID=" & strProductID & "&contentReferenceID=" & paryProductReviews(0, i) & """>Yes</a>" _
							 & "&nbsp;&nbsp;<a href=""" & CurrentPage & "?ReviewAction=RateReviewNo&amp;PRODUCT_ID=" & strProductID & "&contentReferenceID=" & paryProductReviews(0, i) & """>No</a>"
				Response.Write "<br /><a href=""mailTo:contact@mysite.com&subject=offensive review&body=I found this review { " & paryProductReviews(0, i) & "} offensive for the following reasons . . ."">Report this review</a>"
			Next 'i
		End If
		%>
		</td>
	  </tr>
	  
	  <tr>
		<td class="tdContent2" align="center"><a href="<%= pstrURL %>" title="write a review">Write a review</a></td>
	  </tr>
	</table>
	<%
	
End Sub	'writeProductReviewDetails

'***************************************************************************************************************************************************************************************

Sub writeProductReviewSummary(byVal strProductID)

Dim pstrURL

	pstrURL = CurrentPage & "?ReviewAction=WriteReview&amp;PRODUCT_ID=" & strProductID
	Call loadProductReviewSummary(strProductID)

	%>
	<% If Len(mstrErrorMessage_ProductReview) > 0 Then Response.Write "<p>" & mstrErrorMessage_ProductReview & "</p>" %>
	<table class="Section" border="1" cellpadding="0" cellspacing="0">
	  <tr>
		<th class="tdMiddleTopBanner">Product Review</th>
	  </tr>
	  <tr>
		<td class="tdContent2">
		<%
		If mlngNumReviews = 0 Then
			Response.Write "This product has not been rated yet."
		Else
			If mlngNumReviews = 1 Then
				Response.Write "One reviewer gave this a rating of " & ratingDisplayImage(mlngAvgReviewScore)
				Response.Write "<p align=center><a href=""" & CurrentPage & "?ReviewAction=ReadReviews&amp;PRODUCT_ID=" & strProductID & """ title=""Read complete review"">Read review</a></p><br />"
			Else
				Response.Write mlngNumReviews & " reviewers gave this an average rating of " & ratingDisplayImage(mlngAvgReviewScore)
				Response.Write "<p align=center><a href=""" & CurrentPage & "?ReviewAction=ReadReviews&amp;PRODUCT_ID=" & strProductID & """ title=""Read complete reviews"">Read reviews</a></p><br />"
			End If
		End If
		%>
		</td>
	  </tr>
	  <tr>
		<td class="tdContent2" align="center"><a href="<%= pstrURL %>" title="write a review">Write a review</a></td>
	  </tr>
	</table>
	<%
	
End Sub	'writeProductReviewSummary

'***************************************************************************************************************************************************************************************

Function validateReview(byVal strProductID)

Dim pstrResult

	If Len(strProductID) = 0 Or Len(strProductID) > 50 Then
		pstrResult = pstrResult & "<li>Invalid product to review</li>"
	End If

	If Len(contentAuthorName) = 0 Then
		pstrResult = pstrResult & "<li>Please enter your name</li>"
	End If

	If Len(contentAuthorEmail) = 0 Then
		pstrResult = pstrResult & "<li>Please enter your email address</li>"
	End If

	If Len(contentAuthorRating) = 0 Then
		pstrResult = pstrResult & "<li>Please select a rating</li>"
	ElseIf Not isNumeric(contentAuthorRating) Then
		pstrResult = pstrResult & "<li>Invalid rating</li>"
	End If

	validateReview = pstrResult
		
End Function	'validateReview

'***************************************************************************************************************************************************************************************

Sub WriteReviewForm(byVal strProductID)

Dim i
Dim pstrURL

	pstrURL = CurrentPage & "?ReviewAction=SubmitReview&amp;PRODUCT_ID=" & strProductID

%>
<script language="javascript" type="text/javascript">

var whitespace = " \t\n\r";

function isEmpty(s)
{   return ((s == null) || (s.length == 0))
}

function isWhitespace (s)

{   var i;

    // Is s empty?
    if (isEmpty(s)) return true;

    // Search through string's characters one by one
    // until we find a non-whitespace character.
    // When we do, return false; if we don't, return true.

    for (i = 0; i < s.length; i++)
    {   
        // Check that current character isn't whitespace.
        var c = s.charAt(i);

        if (whitespace.indexOf(c) == -1) return false;
    }

    // All characters are whitespace.
    return true;
}

function isEmail (s)
{
   if (isEmpty(s)) 
       if (isEmail.arguments.length == 1) return defaultEmptyOK;
       else return (isEmail.arguments[1] == true);
   
    // is s whitespace?
    if (isWhitespace(s)) return false;
    
    // there must be >= 1 character before @, so we
    // start looking at character position 1 
    // (i.e. second character)
    var i = 1;
    var sLength = s.length;

    // look for @
    while ((i < sLength) && (s.charAt(i) != "@"))
    { i++
    }

    if ((i >= sLength) || (s.charAt(i) != "@")) return false;
    else i += 2;

    // look for .
    while ((i < sLength) && (s.charAt(i) != "."))
    { i++
    }

    // there must be at least one character after the .
    if ((i >= sLength - 1) || (s.charAt(i) != ".")) return false;
    else return true;
}


function validateReviewSumbission(theForm)
{

var blnFound = false;

	if (theForm.contentAuthorName.value == "")
	{
		alert("Please enter your name");
		theForm.contentAuthorName.focus();
		return false;
	}
	
	if (theForm.contentAuthorEmail.value == "")
	{
		alert("Please enter your email address");
		theForm.contentAuthorEmail.focus();
		return false;
	}
	
	if (!isEmail(theForm.contentAuthorEmail.value))
	{
		alert("Please verify your email address. It appears it is in an incorrect format.");
		theForm.contentAuthorEmail.focus();
		return false;
	}
	
	for (var i = 0;  i < theForm.contentAuthorRating.length;  i++)
	{
		if (theForm.contentAuthorRating[i].checked)
		{
			blnFound = true;
		}
	}
	
	if (!blnFound)
	{
		alert("Please select a rating.");
		theForm.contentAuthorRating[theForm.contentAuthorRating.length-1].focus();
		return false;
	}

	return true;
}
</script>
<% If Len(mstrErrorMessage_ProductReview) > 0 Then Response.Write "<p>" & mstrErrorMessage_ProductReview & "</p>" %>
<table class="Section" cellpadding="0" cellspacing="0" border="1">
<tr>
	<th class="tdMiddleTopBanner">Write a review</th>
</tr>
  <tr>
    <td class="tdContent2">        
		<form METHOD="POST" ACTION="<%= pstrURL %>" onSubmit="return validateReviewSumbission(this);" name="frmSubmitReview" ID="frmSubmitReview">
		<center>
		<table cellpadding="4" cellspacing="4" border="0"">
		<tr>
			<td><b>Name:</b></td>
			<td align="left"><input type="text" name="contentAuthorName" id="contentAuthorName" size="40" value="<%= contentAuthorName %>"></td>
		</tr>
		<tr>
			<td><b>Email:</b></td>
			<td align="left"><input type="text" name="contentAuthorEmail" id="contentAuthorEmail" size="40" value="<%= contentAuthorEmail %>"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td align="left"><input type="checkbox" name="contentAuthorShowEmail" id="contentAuthorShowEmail" value="1" <% If contentAuthorShowEmail = "1" Then Response.Write "checked" %>><label for="contentAuthorShowEmail">Display my email address</label></td>
		</tr>
		<tr>
			<td colspan="2">

			<table width="100%" cellpadding="0" cellspacing="0" border="0">
			<tr>
				<td><b>Rating:</b>&nbsp;</td>
				<td><input type="radio" name="contentAuthorRating" id="contentAuthorRating1" value="1" <% If contentAuthorRating = "1" Then Response.Write "checked" %>><label for="contentAuthorRating1"><%= ratingDisplayImage(1) %></label></td>
				<td><input type="radio" name="contentAuthorRating" id="contentAuthorRating2" value="2" <% If contentAuthorRating = "2" Then Response.Write "checked" %>><label for="contentAuthorRating2"><%= ratingDisplayImage(2) %></label></td>
				<td><input type="radio" name="contentAuthorRating" id="contentAuthorRating3" value="3" <% If contentAuthorRating = "3" Then Response.Write "checked" %>><label for="contentAuthorRating3"><%= ratingDisplayImage(3) %></label></td>
				<td><input type="radio" name="contentAuthorRating" id="contentAuthorRating4" value="4" <% If contentAuthorRating = "4" Then Response.Write "checked" %>><label for="contentAuthorRating4"><%= ratingDisplayImage(4) %></label></td>
				<td><input type="radio" name="contentAuthorRating" id="contentAuthorRating5" value="5" <% If contentAuthorRating = "5" Then Response.Write "checked" %>><label for="contentAuthorRating5"><%= ratingDisplayImage(5) %></label></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="2"><b>Message:</b> (Optional)<br />
			<textarea name="contentContent" id="contentContent" cols="50" rows="6"><%= contentContent %></textarea>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="center"><input type="submit" class="btn" value="Submit Review" ID="btnSubmitReview" NAME="btnSubmitReview"></td>
		</tr>
		</table>
		</center>
		</form>
	  </td>
	</tr>
  </table>
<% End Sub	'WriteReviewForm %>