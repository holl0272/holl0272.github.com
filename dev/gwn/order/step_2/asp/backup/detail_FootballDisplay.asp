<%
	Const cbytStartGrapicLogos = 4
	
	Dim mstrJavascript
	Dim paryBaseJerseyImages(5)
	Dim pblnAttributeSelected
	Dim pbytSelectedImagePosition
	Dim plngAttributeToUse
%>	  
<script language="javascript" type="text/javascript">

var selectedJerseyPrintingOption = 0;
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

function showTeamFontWarningMessage(blnDisplay)
{
    if (blnDisplay)
    {
        document.getElementById("divTeamFontWarning").style.display = "";
    }else{
        document.getElementById("divTeamFontWarning").style.display = "none";
    }
}

function viewLargeJersey()
{
    //window.open(aryLargeImage[selectedJerseyColor],'detailImage','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=480,width=640,resizable')
    window.open('displayLargeImage.asp?'+aryLargeImage[selectedJerseyColor],'detailImage','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=480,width=640,resizable')
}

	function setSelectedJerseyImage(index, option)
	{
		var imgTarget;
		var imgSource;

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
			case "2": //number
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
			case "3": //name, number
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
			case "4": //team, number
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
			case "5": //team, name, number
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
			case "6": //team, name, number
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
		selectedJerseyPrintingOption = option;
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

		//alert(strTargetField + ": " + pstrValue);
		//alert(paryTargetField.length);
		
		for (var i = 0;  i < paryTargetField.length;  i++)
		{
		paryTargetField[i].value = pstrValue;
		}

	}

	var blnTeamNameRequired = false;
	var blnPlayerNameRequired = false;
	var blnNumberRequired = false;
	
	function ValidateForm_Jersey(theForm)
	{
        var radLetteringFontOption = document.getElementsByName("dummy_attr<%=  maryJerseyAttributes(enAttrPos_LetteringFont)(enJerseyAttribute_Position) & mstrPrefix %>");
        var radTeamOption = document.getElementsByName("dummy_attr<%=  maryJerseyAttributes(enAttrPos_LetteringStyleTeam)(enJerseyAttribute_Position) & mstrPrefix %>");

		if (selectedJerseyPrintingOption == 4)
		{
		    for (var i = <%= cbytStartGrapicLogos %>;  i < radTeamOption.length;  i++)
		    {
		        if (radTeamOption[i].checked)
		        {
		            if (!radLetteringFontOption[0].checked)
		            {
		                radLetteringFontOption[0].checked = true;
    		            /*alert('bad choice');
                        radLetteringFontOption[0].focus();
                        return false;*/
		            }
		        }
		    }
		}

		if (selectedJerseyPrintingOption > 4)
		{
            if (radLetteringFontOption[0].checked)
            {
	            alert('Please select a lettering font');
                radLetteringFontOption[0].focus();
                return false;
            }
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

	</script>
<form method="post" name="frmSource" id="frmSource" onsubmit="return false;">      
<div id="jerseyColor" class="jerseyOptions">
	<div class="jerseyTitle">Choose Color of Jersey</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<%
	Dim pstrImageArray
	plngAttributeToUse = enAttrPos_JerseyColor
	pblnAttributeSelected = False
	pbytSelectedImagePosition = 0
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	pstrImageArray = "var maryAttrImages = new Array();" & vbcrlf
	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		'Response.Write "<td>" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) And Not pblnAttributeSelected Then
			Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			Response.Write "<script language=""javascript"">selectedJerseyColor = " & i & ";</script>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If
		
		pstrImageArray = pstrImageArray & "maryAttrImages[" & i & "] = " & Chr(34) & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & Chr(34) & ";" & vbcrlf

		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			'Response.Write "<td><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
			Response.Write "<td class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Extra)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td class=""jerseyDisplay"">&nbsp;</td>"
		End If
		
		If i = 0 Then
		    Response.Write "<td rowspan=""" & CStr(UBound(attributeDetail) + 1) & """ style=""text-align:center;border-left:dotted 1pt black;"">" & detailImageOut & "<br /><a href="""" title=""View Larger Image"" onclick=""viewLargeJersey();return false;"">View Large Image</a></td>"
		End If
		Response.Write "</tr>"
	Next 'i

	Response.Write "<script language=""javascript"">var aryLargeImage = new Array();"
	For i = 0 To UBound(attributeDetail)
		Response.Write "aryLargeImage[" & i & "] = '" & attributeDetail(i)(enAttributeDetail_Extra1) & "';"
	Next 'i
	Response.Write "</script>"
	%>	  
	</table>
	<script language="javascript" type="text/javascript"><%= pstrImageArray %></script>
</div>

<div id="jerseyLetteringOptions" class="jerseyOptions">
	<div class="jerseyTitle">Choose Lettering Option</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<%
	plngAttributeToUse = enAttrPos_LetteringOption
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail)
		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">"
		Response.Write attributeDetail(i)(enAttributeDetail_Name)

		Select Case attributeDetail(i)(enAttributeDetail_Type)
			Case 1
			    Response.Write "<br />(Add " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			Case 2
			    Response.Write "<br />(Subtract " & FormatCurrency(CDbl(attributeDetail(i)(enAttributeDetail_Price))) & ")"
			Case Else

		End Select

		Response.Write "</td>"
		Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setJerseyDisplayOptions('" & i & "');setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
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

<div id="jerseyLetteringColor" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Choose Lettering Color</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<%
	plngAttributeToUse = enAttrPos_LetteringColor
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail) Step 2
		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"" checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");""></td>"
		End If
		
		If Len(attributeDetail(i)(enAttributeDetail_Image)) > 0 Then
			Response.Write "<td class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i & """></td>"
		Else
			Response.Write "<td class=""jerseyDisplay"">&nbsp;</td>"
		End If
		
		'now write the 2nd half
		If UBound(attributeDetail) >= i + 1 Then
		    Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i+1)(enAttributeDetail_Name) & "</td>"
		    Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i+1)(enAttributeDetail_Display) & "&nbsp;</td>"
		    If attributeDetail(i+1)(enAttributeDetail_Default) Then
			    Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i+1)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i + 1 & ");"" checked></td>"
			    paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i+1)(enAttributeDetail_Image)
			    pblnAttributeSelected = True
		    Else
			    Response.Write "<td class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ ID=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i+1)(enAttributeDetail_ID) & """ onclick=""setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i + 1 & ");""></td>"
		    End If
		
		    If Len(attributeDetail(i+1)(enAttributeDetail_Image)) > 0 Then
			    Response.Write "<td class=""jerseyDisplay""><img border=""0"" src=""" & Server.HTMLEncode(attributeDetail(i+1)(enAttributeDetail_Image)) & """ alt=""" & attributeDetail(i+1)(enAttributeDetail_Name) & """ id=""imgJerseyOption_" & pbytSelectedImagePosition & "_" & i + 1 & """></td>"
		    Else
			    Response.Write "<td class=""jerseyDisplay"">&nbsp;</td>"
		    End If
		Else
			Response.Write "<td colspan=""3"" class=""jerseyDisplay"">&nbsp;</td>"
		End If
		
		Response.Write "</tr>"
	Next 'i
	%>	  
	</table>
</div>

<div id="jerseyLetteringFont" class="jerseyOptions" style="display:none">
	<div class="jerseyTitle">Choose Lettering Font for Team and Player Names</div>
	<table border="0" cellpadding="2" cellspacing="0" class="jerseyDisplay">
	<colgroup>
	  <col width="50%" />
	  <col width="10%" />
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
		Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
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
	  <col width="50%" />
	  <col width="10%" />
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
		Response.Write "<td  class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
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
	  <col width="50%" />
	  <col width="10%" />
	  <col width="40%" />
	</colgroup>
	<tr><td class="jerseyTitle" colspan="4">Letter Only Styles</td></tr>
	<%
	plngAttributeToUse = enAttrPos_LetteringStyleTeam
	pblnAttributeSelected = False
	pbytSelectedImagePosition = pbytSelectedImagePosition + 1
	attributeDetail = maryAttributes(maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) - 1)(enAttribute_DetailArray)	'need to subtract one to adjust for attribute array
	For i = 0 To UBound(attributeDetail)

        'add the graphics message
        If i = cbytStartGrapicLogos Then
		    Response.Write "<tr><td class=""jerseyTitle"" colspan=""4"">Graphics Styles</td></tr>"
        End If
        
        If i >= cbytStartGrapicLogos And i <> UBound(attributeDetail) Then
            mstrJavascript = """showTeamFontWarningMessage(true);setCustomLogo(" & i & "," & UBound(attributeDetail) & ");setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"""
        Else
            mstrJavascript = """showTeamFontWarningMessage(false);setCustomLogo(" & i & "," & UBound(attributeDetail) & ");setTeamValue(this,'attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & "'," & pbytSelectedImagePosition & "," & i & ");"""
        End If
        
		Response.Write "<tr>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Name) & "</td>"
		Response.Write "<td class=""jerseyDisplay"">" & attributeDetail(i)(enAttributeDetail_Display) & "&nbsp;</td>"
		If attributeDetail(i)(enAttributeDetail_Default) Then
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ id=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=" & mstrJavascript & " checked></td>"
			paryBaseJerseyImages(pbytSelectedImagePosition) = attributeDetail(i)(enAttributeDetail_Image)
			pblnAttributeSelected = True
		Else
			Response.Write "<td  class=""jerseyDisplay""><input type=""radio"" name=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ id=""dummy_attr" & maryJerseyAttributes(plngAttributeToUse)(enJerseyAttribute_Position) & mstrPrefix & """ value=""" & attributeDetail(i)(enAttributeDetail_ID) & """ onclick=" & mstrJavascript & "></td>"
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
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(2) %>" id="imgSelectedJersey2"><br /><img border="0" src="<%= paryBaseJerseyImages(3) %>" id="imgSelectedJersey3"></td>
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(4) %>" id="imgSelectedJersey4"></td>
      <td align="center" valign="top"><img border="0" src="<%= paryBaseJerseyImages(5) %>" id="imgSelectedJersey5"></td>
    </tr>
  </table>
  <div id="divTeamFontWarning" style="display:none">
  <%= getPageFragmentByKey("TeamFontWarning") %>
  </div>
</div>
</form>
