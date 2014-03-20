<%
	Dim paryBaseJerseyImages(6)
	Dim pblnAttributeSelected
	Dim pbytSelectedImagePosition
	Dim plngAttributeToUse
%>
<script language="javascript">

var selectedJerseyColor;

function setCustomLogo(bytOption, bytNumOptions)
{
    if (bytOption == (bytNumOptions-1))
    {
        document.getElementById("customLogo").value = 1;
    }else{
        document.getElementById("customLogo").value = "";
    }
}

function viewLargeJersey()
{
    window.open('displayLargeImage.asp?'+aryLargeImage[selectedJerseyColor],'detailImage','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=480,width=640,resizable')
}

function viewLargeJersey2ndSide()
{
    window.open(aryLargeImage2nd[selectedJerseyColor],'detailImage','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=480,width=640,resizable')
}

	function setSelectedJerseyImage(index, option)
	{
		var imgTarget;
		var imgSource;

        if (index == undefined) return false;
        if (index.length == 0)
        {
            return false;
        }

        if (index == 0)
        {
            selectedJerseyColor = option;
		    imgSource = maryAttrImages[option];
 		    imgTarget = document.getElementById("imgSelectedJersey"+index);
 		    imgTarget.src = imgSource;
 		    imgTarget = document.getElementById("imgDetail");
 		    imgTarget.src = imgSource;
        }else{
		    imgTarget = document.getElementById("imgSelectedJersey"+index);
		    imgSource = document.getElementById("imgJerseyOption_"+index+"_"+option);
		    if (imgSource)
		    {
			    imgTarget.src = imgSource.src;
		    }else{
			    if (imgTarget) imgTarget.src = "images/transparent.gif";
		    }
        }
	}

	function setJerseyDisplayOptions(option)
	{
	/*
	jerseyLetteringColor
	jerseyLetteringFont
	jerseyPlayerOptions
	jerseyTeamOptions
	*/

	var styleColName;
	var styleColNumber;

		if (document.styleSheets[1].rules)
		{
			styleColNumber = document.styleSheets[1].rules.item(1).style;
			styleColName =document.styleSheets[1].rules.item(0).style;
		}
		else if (document.styleSheets[1].cssRules)
		{
			styleColNumber = document.styleSheets[1].cssRules.item(1).style;
			styleColName =document.styleSheets[1].cssRules.item(0).style;
		}

		switch (option)
		{
			case "0": //no printing
				SidesSelected = 0;
				hideElement(document.getElementById("jerseyLetteringColor"));
				hideElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "none";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = false;
				break;
			case "1": //numbers
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				hideElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "2": //numbers
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				hideElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "3": //number
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				hideElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "4": //number
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				hideElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "5": //name, number
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			case "6": //name, number
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				hideElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			case "7": //team, number
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "8": //team, number
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				hideElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = false;
				blnNumberRequired = true;
				break;
			case "9": //team, name, number
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			case "10": //team, name, number
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			case "11": //team, name, number
				SidesSelected = 1;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			case "12": //team, name, number
				SidesSelected = 2;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "";
				styleColNumber.display = "";
				blnTeamNameRequired = true;
				blnPlayerNameRequired = true;
				blnNumberRequired = true;
				break;
			default:
				SidesSelected = 0;
				showElement(document.getElementById("jerseyLetteringColor"));
				showElement(document.getElementById("jerseyLetteringFont"));
				showElement(document.getElementById("jerseyPlayerOptions"));
				showElement(document.getElementById("jerseyTeamOptions"));
				styleColName.display = "none";
				styleColNumber.display = "none";
				blnTeamNameRequired = false;
				blnPlayerNameRequired = false;
				blnNumberRequired = false;
				break;
		}
	}

	function set2ndColorImage(index, option)
	{
	    var imageIndex = index + 1;

	    imgTarget = document.getElementById("imgSelectedJersey"+imageIndex);
	    imgSource = document.getElementById("imgJerseyOption_"+index+"_"+option);
	    if (imgSource)
	    {
		    imgTarget.src = imgSource.src;
	    }else{
		    if (imgTarget) imgTarget.src = "images/transparent.gif";
	    }

		//setSelectedJerseyImage(bytDisplayImage, bytOptionIndex)
    }

	function setSmallReversibleImage(strSrc)
	{
	    var src;

	    if (strSrc.length == 0)
	    {
	        src = "images/transparent.gif"
	    }else{
	        src = strSrc
	    }

	    imgTarget = document.getElementById("imgSmallReversibleImage");
	    if (imgTarget) imgTarget.src = src;
    }

	function setTeamValue(theSourceField, strTargetField, bytDisplayImage, bytOptionIndex)
	{

	var pstrValue = theSourceField.value;
	var paryTargetField = eval("document.frmDetail." + strTargetField);

		setSelectedJerseyImage(bytDisplayImage, bytOptionIndex)
		if (pstrValue == null)
		{
			pstrValue = theSourceField.options[theSourceField.selectedIndex].value
		}

		// alert(strTargetField + ": " + pstrValue);
		// alert(paryTargetField.length);

		for (var i = 0;  i < paryTargetField.length;  i++)
		{
		paryTargetField[i].value = pstrValue;
		}

	}

	var blnTeamNameRequired = false;
	var blnPlayerNameRequired = false;
	var blnNumberRequired = false;
	var SidesSelected = 0;

	function ValidateForm_Jersey(theForm)
	{
		if (!validateSelections())
		{
		    return false;
		}

		if (blnTeamNameRequired)
		{
	        var theTeamName = document.getElementById("dummy_attr<%= maryJerseyAttributes(enAttrPos_TeamName)(enJerseyAttribute_Position) & mstrPrefix %>");
		    if (theTeamName.value == "")
		    {
		        alert("Please enter a team name.");
		        theTeamName.focus();
		        return false;
		    }
		}

		if (!checkQuantityOrdered(theForm))
		{
			theForm.QUANTITY[0].focus();
			alert("Please select a quantity.");
			return false;
		}

		return validateLineItems(theForm);

	}

	function checkQuantityOrdered(theForm)
	{
	var theQty = theForm.QUANTITY;

		for (var i = 0;  i < theQty.length;  i++)
		{
			if (theQty[i].value != ""){return true;}
		}

		return false;

	}

	function validateSelections()
	{
	    var radLetteringOption = frmSource.dummy_attr<%=  maryJerseyAttributes(enAttrPos_LetteringOption)(enJerseyAttribute_Position) & mstrPrefix %>;
	    if (radLetteringOption[0].checked) return true;

	    var radLetterColorSide1 = frmSource.dummy_attr<%= maryJerseyAttributes(enAttrPos_LetteringColor)(enJerseyAttribute_Position) & mstrPrefix %>;
	    var radLetterColorSide2 = frmSource.dummy_attr<%= maryJerseyAttributes(enAttrPos_LetteringColor_Side2)(enJerseyAttribute_Position) & mstrPrefix %>;

		switch (SidesSelected)
		{
			case 0:
				if (!radLetterColorSide1[0].checked)
				{
					alert("To select a lettering color you must first select a lettering option.");
					radLetterColorSide1[0].focus();
					return false;
				}
				if (!radLetterColorSide2[0].checked)
				{
					alert("To select a lettering color you must first select a lettering option.");
					radLetterColorSide2[0].focus();
					return false;
				}
				break;

			case 1:
				if ((radLetterColorSide1[0].checked) && (radLetterColorSide2[0].checked))
				{
					alert("Please select a lettering color.");
					radLetterColorSide1[0].focus();
					return false;
				}

				if (!radLetterColorSide1[0].checked && !radLetterColorSide2[0].checked)
				{
					alert("For the option chosen, please select only one lettering color.");
					radLetterColorSide2[0].focus();
					return false;
				}
				break;
			case 2:
				if (radLetterColorSide1[0].checked)
				{
					alert("For the option chosen, please select two lettering colors.");
					radLetterColorSide1[0].focus();
					return false;
				}

				if (radLetterColorSide2[0].checked)
				{
					alert("For the option chosen, please select two lettering colors.");
					radLetterColorSide2[0].focus();
					return false;
				}
				break;
			default:
		}

	    return true;
    }

	function validateLineItems(theForm)
	{
	    var theQty = theForm.QUANTITY;
	    var theNumber = theForm.attr<%= maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_Number)(enJerseyAttribute_Default) %>;
	    var theName = theForm.attr<%= maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_NameOnJersey)(enJerseyAttribute_Default) %>;

		for (var i = 0;  i < theQty.length;  i++)
		{
			if (theQty[i].value != "")
			{
			    if (blnNumberRequired)
			    {
			        if (theNumber[i].value == "")
			        {
			            alert("Please enter a number.");
			            theNumber[i].focus();
			            return false;
			        }
			    }

			    if (blnPlayerNameRequired)
			    {
			        if (theName[i].value == "")
			        {
			            alert("Please enter a name.");
			            theName[i].focus();
			            return false;
			        }
			    }

			}else{
			    if ((theName[i].value + theNumber[i].value) != "")
			    {
			        alert("Please enter a quantity.");
			        theQty[i].focus();
			        return false;
			    }
			}
		}

		return true;
	}

	function setColors(Color1, Color2)
	{
	    document.getElementById("spSide1Color").innerHTML = Color1;
	    document.getElementById("spSide2Color").innerHTML = Color2;
	    document.getElementById("spSide1Color_Selected").innerHTML = Color1;
	    document.getElementById("spSide2Color_Selected").innerHTML = Color2;
    }
	</script>
<form method="post" name="frmSource" ID="frmSource" onsubmit="return false;">
<div id="jerseyColor" class="jerseyOptions">
	<div class="jerseyTitle">Choose Color of Jersey</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<%
	Dim pstrImageArray
	Dim pstrJavascript
	Dim pstrColor1
	Dim pstrColor2
	Dim paryColor
	Dim SelectedSide1Color
	Dim SelectedSide2Color
	Dim SmallReversibleImage

    SmallReversibleImage = "images/transparent.gif"
	plngAttributeToUse = enAttrPos_JerseyColor
	pblnAttributeSelected = False
	pbytSelectedImagePosition = 0
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	pstrImageArray = "var maryAttrImages = new Array();var maryAttrImages2nd = new Array();" & vbcrlf
	For i = 0 To UBound(attributeDetail)
	    paryColor = Split(attributeDetail(i)(enAttributeDetail_Name), "/")
	    If UBound(paryColor) >=0 Then pstrColor1 = paryColor(0)
	    If UBound(paryColor) >=0 Then pstrColor2 = paryColor(1)
        pstrJavascript = " onclick=""setColors('" & pstrColor1 & "', '" & pstrColor2 & "');setSmallReversibleImage('" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Display)) & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"
		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		'Response.Write "<td>" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) And Not pblnAttributeSelected Then
			Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & Chr(34) & pstrJavascript & """ checked></td>"
			Response.Write "<script language=""javascript"">selectedJerseyColor = " & i & ";</script>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
			SelectedSide1Color = pstrColor1
			SelectedSide2Color = pstrColor2
		    SmallReversibleImage = Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Display))
		Else
			Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & Chr(34) & pstrJavascript & """></td>"
		End If

		pstrImageArray = pstrImageArray & "maryAttrImages[" & i & "] = " & Chr(34) & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & Chr(34) & ";" & vbcrlf
		pstrImageArray = pstrImageArray & "maryAttrImages2nd[" & i & "] = " & Chr(34) & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Extra)) & Chr(34) & ";" & vbcrlf

		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			'Response.Write "<td><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
			Response.Write "<td class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Display)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td class=""jerseyDisplay"">&nbsp;</td>"
		End If

		If i = 0 Then
		    Response.Write "<td rowspan=""" & CStr(UBound(attributeDetail) + 1) & """ style=""text-align:center;border-left:dotted 1pt black;"">" & detailImageOut & "<br /><a href="""" title=""View Larger Image"" onclick=""viewLargeJersey();return false;"">View Large Image of Both Sides of Jersey</a></td>"
		    'Response.Write "<td rowspan=""" & CStr(UBound(attributeDetail) + 1) & """ style=""text-align:center;border-left:dotted 1pt black;"">" & detailImageOut & "<br /><a href="""" title=""View Larger Image"" onclick=""viewLargeJersey();return false;"">View Large Image</a> | <a href="""" title=""View 2nd Side"" onclick=""viewLargeJersey2ndSide();return false;"">View Large Image 2nd</a></td>"
		End If
		Response.Write "</tr>"
	Next 'i

	Response.Write "<script language=""javascript"">var aryLargeImage = new Array();var aryLargeImage2nd = new Array();"
	For i = 0 To UBound(attributeDetail)
		Response.Write "aryLargeImage[" & i & "] = '" & attributeDetail(i)(enAttributeDetail_Extra1) & "';"
		Response.Write "aryLargeImage2nd[" & i & "] = '" & attributeDetail(i)(enAttributeDetail_Extra) & "';"
	Next 'i
	Response.Write "</script>"
	%>
	</table>
	<script language="javascript" type="text/javascript"><%= pstrImageArray %></script>
</div>

<div id="jerseyLetteringOptions" class="jerseyOptions">
	<div class="jerseyTitle">Choose Lettering Option</div>
	<table border="1" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<%
	plngAttributeToUse = enAttrPos_LetteringOption
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array

	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"
		Response.Write "<td align=""left"">"

		'write the text
		Response.Write "<table border=""0"" style=""width:100%;"">"
		Response.Write "<colgroup><col width=""75%"" /><col width=""25%"" /></colgroup>"
		If Len(attributeDetail(i)(enAttributeDetail_Display)) > 0 Then Response.Write "<tr><td colspan=""2"" class=""jerseyDisplay2"">" & attributeDetail(i)(enAttributeDetail_Display) & "</td></tr>"

		'write the attribute
		Response.Write "<tr>"
		Response.Write "<td valign=""bottom"" class=""jerseyDisplay2"">" & attributeDetail(i)(enAttributeDetail_Name) & ": "
		Select Case attributeDetail(i)(enAttributeDetail_Type)
			Case 1
			    Response.Write "(Add " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			Case 2
			    Response.Write "(Subtract " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			Case Else

		End Select
		Response.Write "<td class=""jerseyDisplay2"">"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"">"
		End If
		Response.Write "</td>"

		'write the reversible attribute
		If i <> 0 Then
		    i = i + 1
		    Response.Write "<tr>"
		    Response.Write "<td valign=""bottom"" class=""jerseyDisplay2"">" & attributeDetail(i)(enAttributeDetail_Name) & ": "
		    Select Case attributeDetail(i)(enAttributeDetail_Type)
			    Case 1
			        Response.Write "(Add " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			    Case 2
			        Response.Write "(Subtract " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			    Case Else

		    End Select
		    Response.Write "<td valign=""bottom"" class=""jerseyDisplay2"">"
		    If attributeDetail(i)(enAttributeDetail_Default) Then
			    Response.Write "<input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i-1 & ");"" checked>"
			    paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			    pblnAttributeSelected = True
		    Else
			    Response.Write "<input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i-1 & ");"">"
		    End If
		    Response.Write "</td>"
		    If Len(attributeDetail(i)(enAttributeDetail_Display)) > 0 Then Response.Write "<tr><td colspan=""2"" class=""jerseyDisplay2"">" & attributeDetail(i)(enAttributeDetail_Display) & "</td></tr>"
		    Response.Write "</td>"
		End If  'i <> 0
		Response.Write "</tr>"
		Response.Write "</table>"


		'write the image
		Response.Write "<td>"
        If i > 0 Then
		    If Len(attributeDetail(i-1)(enAttributeDetail_Image)) > 0 Then
			    Response.Write "<img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i-1)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i-1)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i-1 & """><br />"
		    End If
        Else
 		    If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			    Response.Write "<img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """><br />"
		    End If
       End If
		Response.Write "</td></tr>"

	Next 'i
	%>
	</table>
</div>

<div id="jerseyLetteringColor" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Choose Lettering Color</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<tr style="border-bottom:solid 1pt black">
	  <td style="border-bottom:dotted 1pt black; border-right: solid 1pt black;">&nbsp;</td>
	  <td style="text-align:center;border-bottom:solid 1pt black">Color on <span id="spSide1Color"><%= SelectedSide1Color %></span> Side</td>
	  <td style="text-align:center;border-bottom:solid 1pt black; border-left: solid 1pt black;"><img border="0" src="<%= SmallReversibleImage %>" id="imgSmallReversibleImage"></td>
	  <td style="text-align:center;border-bottom:solid 1pt black; border-left: solid 1pt black;">Color on <span id="spSide2Color"><%= SelectedSide2Color %></span> Side</td>
	  <td style="border-bottom:solid 1pt black">&nbsp;</td>
	</tr>
	<%
	plngAttributeToUse = enAttrPos_LetteringColor
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1

	'get the two color attributes
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	attributeDetail2 = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position))(enAttribute_DetailArray)
	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"

		'write the attribute detail name
		Response.Write "<td class=""jerseyDisplay"" style=""border-right: solid 1pt black;"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"

		'write the first side radio button
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay"" style=""text-align:center;""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td class=""jerseyDisplay"" style=""text-align:center;""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If

		'write the center image
		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			Response.Write "<td class=""jerseyDisplay"" style=""text-align:center;border-left:solid 1pt black""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td class=""jerseyDisplay"" style=""text-align:center;border-left:solid 1pt black"">&nbsp;</td>"
		End If

		'write the second side radio button
		If attributeDetail(i)(enAttributeDetail_Default) Then
		    Response.Write "<td class=""jerseyDisplay"" style=""text-align:center;border-left: solid 1pt black;""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail2(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & "','','');set2ndColorImage(" & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
		Else
		    Response.Write "<td class=""jerseyDisplay"" style=""text-align:center;border-left: solid 1pt black;""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail2(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse+1)(enJerseyAttribute_Position) & mstrPrefix & "','','');set2ndColorImage(" & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If

		Response.Write "</tr>"
	Next 'i

	'need to increment selected image position by one more since 2nd side color is already covered
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	%>
	</table>
</div>

<div id="jerseyLetteringFont" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Lettering Font for Team and Player Names</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<colgroup>
	  <col width="40%" />
	  <col width="20%" />
	  <col width="40%" />
	</colgroup>
	<%
	plngAttributeToUse = enAttrPos_LetteringFont
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"
		Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		'Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If

		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			Response.Write "<td  class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td  class=""jerseyDisplay"">&nbsp;</td>"
		End If
		Response.Write "</tr>"
	Next 'i
	%>
	</table>
	<div style="text-align:center; padding-top:12pt; padding-bottom:12pt">
	  Number Font is Pro Narrow<br />
	  Not applicable for Team Name if you choose "Letters with Graphics" below
	</div>
</div>

<div id="jerseyPlayerOptions" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Player Name Lettering Style</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<colgroup>
	  <col width="40%" />
	  <col width="20%" />
	  <col width="40%" />
	</colgroup>
	<%
	plngAttributeToUse = enAttrPos_LetteringStyleName
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"
		Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		'Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If

		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			Response.Write "<td  class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td  class=""jerseyDisplay"">&nbsp;</td>"
		End If
		Response.Write "</tr>"
	Next 'i
	%>
	</table>
</div>

<div id="jerseyTeamOptions" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Team Name Lettering Style</div>

	<div class="jerseyDisplay">
	Team Name: <input type="text" name="dummy_attr<%= maryJerseyAttributes(enAttrPos_TeamName)(enJerseyAttribute_Position) & mstrPrefix %>" ID="dummy_attr<%= maryJerseyAttributes(enAttrPos_TeamName)(enJerseyAttribute_Position) & mstrPrefix %>" onchange="setTeamValue(this,'attr<%= maryJerseyAttributes(enAttrPos_TeamName)(enJerseyAttribute_Position) & mstrPrefix & cstrSSTextBasedAttributeHTMLDelimiter & maryJerseyAttributes(enAttrPos_TeamName)(enJerseyAttribute_Default) %>')" size="20"><br />
	Location of Team Name:
	<select name="dummy_attr<%= maryJerseyAttributes(enAttrPos_LocationTeam)(enJerseyAttribute_Position) & mstrPrefix %>" ID="dummy_attr<%= maryJerseyAttributes(enAttrPos_LocationTeam)(enJerseyAttribute_Position) & mstrPrefix %>" size="1" onchange="setTeamValue(this,'attr<%= maryJerseyAttributes(enAttrPos_LocationTeam)(enJerseyAttribute_Position) & mstrPrefix %>');">
	<%
		attributeDetail = maryAttributes(maryJerseyAttributes(enAttrPos_LocationTeam)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
		For i = 0 To UBound(attributeDetail)
			If attributeDetail(i)(enAttributeDetail_Default) Then
				Response.Write "<option value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ selected>" & attributeDetail(i)(enAttributeDetail_Name) & "</option>"
			Else
				Response.Write "<option value=""" & attributeDetail(i)(enAttributeDetail_ID) & """>" & attributeDetail(i)(enAttributeDetail_Name) & "</option>"
			End If
		Next 'i
	%>
	</select>
	</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<colgroup>
	  <col width="40%" />
	  <col width="20%" />
	  <col width="40%" />
	</colgroup>
	<tr><td class="jerseyTitle" colspan="4">Letters Only</td></tr>
	<%
	plngAttributeToUse = enAttrPos_LetteringStyleTeam
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail)

        'add the graphics message
        If i = 4 Then
		    Response.Write "<tr><td class=""jerseyTitle"" colspan=""3"">Letters with Graphics</td></tr>"
        End If

		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		'Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ id=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setCustomLogo(" & i & "," & UBound(attributeDetail) & ");setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ id=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setCustomLogo(" & i & "," & UBound(attributeDetail) & ");setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If

		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			Response.Write "<td  class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td  class=""jerseyDisplay"">&nbsp;</td>"
		End If
		Response.Write "</tr>"
	Next 'i
	%>
	</table>

</div>

<%
For i = 0 To UBound(paryBaseJerseyImages)
	If Len(paryBaseJerseyImages(i)) = 0 Then paryBaseJerseyImages(i) = "images/transparent.gif"
Next 'i
%>
<div id="jerseySelectedOptions" class="jerseyOptions">
  <table id="tblSelectedJersey" border="1" cellpadding="2" cellspacing="0" class="jerseyDisplay">
    <tr><th colspan="6" class="jerseyTitle">Selected Jersey</th></tr>
    <tr>
      <th align="center" valign="top">Jersey</th>
      <th align="center" valign="top">Lettering Style<div style="font-size:8pt; font-variant:normal;">(if applicable)</div></th>
      <th align="center" valign="top">Player Name Lettering Style<div style="font-size:8pt; font-variant:normal;">(if applicable)</div></th>
      <th align="center" valign="top">Team Name Lettering Style<div style="font-size:8pt; font-variant:normal;">(if applicable)</div></th>
    </tr>
    <tr>
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(0) %>" id="imgSelectedJersey0"><br /><img border="0" src="<%= paryBaseJerseyImages(1) %>" id="imgSelectedJersey1"></td>
      <td align="center" valign="top" nowrap="nowrap">
        <span id="spSide1Color_Selected"><%= SelectedSide1Color %></span> Side: <img border="0" src="<%= paryBaseJerseyImages(2) %>" id="imgSelectedJersey2"><br />
        <span id="spSide2Color_Selected"><%= SelectedSide2Color %></span> Side: <img border="0" src="<%= paryBaseJerseyImages(2) %>" id="imgSelectedJersey3"><br />

        <img border="0" src="<%= paryBaseJerseyImages(3) %>" id="imgSelectedJersey4">
      </td>
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(4) %>" id="imgSelectedJersey5"></td>
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(5) %>" id="imgSelectedJersey6"></td>
    </tr>
  </table>
</div>
</form>
