<%
'********************************************************************************
'*   Promotion Manager for StoreFront 5.0										*
'*   Release Version:	2.00.001 												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		N/A														*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

Dim cblnIsSQL
cblnIsSQL = CBool(Application("AppDatabase") <> "Access")

								'Set this value to true for SQL databases
'cblnIsSQL = False				'Only need to set this manually for very early versions

dim mrsPromotions

	Call LoadPromotions
	Call DisplayPromotions

On Error Resume Next

	mrsPromotions.Close
	set mrsPromotions = nothing
	cnn.Close
	set cnn = nothing

'************************************************************************************************

Private Function LoadPromotions

dim sql

On Error Resume Next

	If cblnIsSQL Then
		sql = "SELECT * " _
			& "FROM Promotions " _
			& "WHERE (((Promotions.StartDate)<='" & Date() & "' Or (Promotions.StartDate) Is Null) AND ((Promotions.EndDate)>='" & Date() & "' Or (Promotions.EndDate) Is Null) AND (Promotions.Inactive=0) AND ((Promotions.Duration) Is Null Or ((Promotions.Duration)>=Convert(int,[StartDate]-'" & Date() & "'))) AND ((Promotions.MaxUses) Is Null Or (Promotions.NumUses) Is Null Or ((Promotions.MaxUses)>[NumUses])))"
'			& "WHERE (((Promotions.StartDate)<='" & Date() & "' Or (Promotions.StartDate) Is Null) AND ((Promotions.EndDate)>='" & Date() & "' Or (Promotions.EndDate) Is Null) AND (Promotions.Inactive=0) AND ((Promotions.Duration) Is Null Or ((Promotions.Duration)>=[StartDate]-'" & Date() & "')) AND ((Promotions.MaxUses) Is Null Or (Promotions.NumUses) Is Null Or ((Promotions.MaxUses)>[NumUses])))"
	Else
		sql = "SELECT * " _
			& "FROM Promotions " _
			& "WHERE (((Promotions.StartDate)<=Date() Or (Promotions.StartDate) Is Null) AND ((Promotions.EndDate)>=Date() Or (Promotions.EndDate) Is Null) AND ((Promotions.Inactive)=False) AND ((Promotions.Duration) Is Null Or ((Promotions.Duration)>=[StartDate]-Date())) AND ((Promotions.MaxUses) Is Null Or (Promotions.NumUses) Is Null Or ((Promotions.MaxUses)>[NumUses])))"
	End If

	Set mrsPromotions = CreateObject("adodb.recordset")
	with mrsPromotions
		.ActiveConnection = cnn
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = sql
		.Open
	end with
	
	If Err.number = 0 Then
		LoadPromotions = (mrsPromotions.RecordCount > 0)
	Else
		Response.Write "<h3 color=red>Error " & Err.number & ": " & Err.Description & "</h3>"
		Response.Write "<h3 color=red>sql = " & sql & "</h3>"
		Err.Clear
		LoadPromotions = False
	End If
	
End Function	'LoadPromotions

Sub DisplayPromotions
%>

<TABLE border=0 cellPadding=1 cellSpacing=1 width="75%">
  
<%
dim pProductID
dim pstrPromotion
dim pstrPromotionDiscount
dim pstrExpires
dim pMinSubTotal
dim j
dim pRSProduct
dim aProducts
dim sqlWhere

for i=1 to mrsPromotions.RecordCount

	pstrPromotion = ""
	pstrPromotionDiscount = ""
	
	pProductID = trim(mrsPromotions("ProductID").value & "")
	pMinSubTotal = CDbl(mrsPromotions("MinSubTotal").value)
	if mrsPromotions("Percentage").value then
		pstrPromotionDiscount = mrsPromotions("Discount").value & "%"
	else
		pstrPromotionDiscount = "$" & mrsPromotions("Discount").value
	end if
	
	if len(pProductID) = 0 then
		if pMinSubTotal = 0 then
			pstrPromotion = "<p>" & pstrPromotionDiscount & " off your purchase. </p>"
		else
			pstrPromotion = "<p>" & pstrPromotionDiscount & " off your purchase with a minimum order $" & pMinSubTotal & ".</p>"
		end if
	else
	
		aProducts = split(pProductID,";")
		sqlWhere = " where prodID='" & aProducts(1) & "'"
		for j = 2 to (ubound(aProducts) - 1)
			sqlWhere = sqlWhere & " or prodID='" & aProducts(j) & "'"
		next 'j
		Set prsProduct = CreateObject("adodb.recordset")
		with prsProduct
			.ActiveConnection = cnn
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = "Select prodID, prodNamePlural from sfProducts " & sqlWhere
			.Open
		end with
		
		if pRSProduct.RecordCount > 0 then
			for j=1 to pRSProduct.RecordCount
				if pMinSubTotal = 0 then
					pstrPromotion = pstrPromotion & "<p>" & pstrPromotionDiscount & " off <a href='Detail.asp?Product_ID=" _
								  & prsProduct("prodID") & "'>" & prsProduct("prodNamePlural") & "</a></p>" & vbcrlf
				else
					pstrPromotion = pstrPromotion & "<p>" & pstrPromotionDiscount & " off <a href='Detail.asp?Product_ID=" _
								  & prsProduct("prodID") & "'>" & prsProduct("prodNamePlural") & "</a> with a minimum order of $" & pMinSubTotal & ".</p>" & vbcrlf
				end if
				prsProduct.movenext
			next 'j
		end if
	end if

	if len(mrsPromotions("EndDate").value & "") > 0 then
		pstrExpires = "&nbsp;&nbsp;&nbsp;Expires on:&nbsp;" & FormatDateTime(mrsPromotions("EndDate"),0)
	elseif len(mrsPromotions("Duration").value & "") > 0 then
		pstrExpires = "&nbsp;&nbsp;&nbsp;Expires on:&nbsp;" & FormatDateTime(mrsPromotions("StartDate")+mrsPromotions("Duration"),0)
	elseif mrsPromotions("MaxUses") > 0 then
		pstrExpires = "<i>Only available for the first " & mrsPromotions("MaxUses") & " customers.</i>"
	end if
%>
  <TR>
    <TD colspan=2>
      <a href="PromotionRegistration.asp?PromoCode=<%= trim(mrsPromotions("PromoCode")) %>"><%= trim(mrsPromotions("PromoTitle").value) %></a>
    </TD>
  </TR>
  <TR>
	<TD Width=10%>&nbsp;</TD>
    <TD>
	  <%= pstrPromotion %>
    </TD>
  </TR>
<% if not (len(mrsPromotions("PromoRules").value & "") = 0) then %>
  <TR>
	<TD Width=10%>&nbsp;</TD>
    <TD>
	  <%= mrsPromotions("PromoRules").value %>
    </TD>
  </TR>
<% end if %>
<% if len(pstrExpires) > 0 then %>
  <TR>
	<TD Width=10%>&nbsp;</TD>
    <TD>
	  <%= pstrExpires %>
    </TD>
  </TR>
<% end if %>
<%
	mrsPromotions.MoveNext
next 'i
%>
</TABLE>
<%

On Error Resume Next

pRSProduct.Close
Set pRSProduct = Nothing

End	Sub	'DisplayPromotions
%>



