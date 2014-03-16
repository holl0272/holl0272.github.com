//********************************************************************************
//*   Common Support File			                                             *
//*   Release Version:   2.00					                                 *
//*   Release Date:      July 4, 2002			                                 *
//*                                                                              *
//*                                                                              *
//*   The contents of this file are protected by United States copyright laws    *
//*   as an unpublished work. No part of this file may be used or disclosed      *
//*   without the express written permission of Sandshot Software.               *
//*                                                                              *
//*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
//********************************************************************************


function getRadio(theRadio)
{
	if (theRadio.length > 0)
	{
		for (var i = 0;  i < theRadio.length;  i++)
		{
			if (theRadio[i].checked)
			{
			return(theRadio[i].value);
			}
		}
		return('');
	}else{
		if (theRadio.checked)
		{
			return(theRadio.value);
		}else{
			return('');
		}
	}
}

function SetRadio(theRadio,theValue)
{

 for (var i=0; i < theRadio.length;i++)
 {
	if (theRadio[i].value == theValue)
	{
		theRadio[i].checked = true;
	}
 }	

}

function SetSelect(theSelect,theValue)
{

 for (var i=0; i < theSelect.length;i++)
 {
	if (theSelect.options[i].value == theValue)
	{
		theSelect.options[i].selected = true;
	}
 }	

}

function OpenHelp(strURL)
{
window.open(strURL,"OrderHelp","toolbar=0,location=0,directories=0,status=0,copyhistory=0,scrollbars=1");
}

function isInteger(theField, emptyOK, theMessage)
{
  if (theField.value == "")
  {
	if (emptyOK)
	{
		return(true);
	}
	{
		alert(theMessage);
		theField.focus();
		theField.select();
	  return (false);
	}
  }

    var i;
    var s = theField.value;
    for (i = 0; i < s.length; i++)
    {
        var c = s.charAt(i);
        if (!((c >= "0") && (c <= "9")))
        {
			alert(theMessage);
			theField.focus();
			theField.select();
            return (false);
        }
    }

  return (true);
}

function isNumeric(theField, emptyOK, theMessage)
{
  if (theField.value == "")
  {
	if (emptyOK)
	{
		return(true);
	}
	{
		alert(theMessage);
		theField.focus();
		theField.select();
	  return (false);
	}
  }
  
  var checkOK = "0123456789.";
  var checkStr = theField.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (allValid)
  { return(true) }
  {
	alert(theMessage);
	theField.focus();
	theField.select();
    return (false);
  }
}

function isEmpty(theField,theMessage)
{
if (theField.value == "")
	{	
	alert(theMessage);
	theField.focus();
	theField.select();
	return(true);
	}
	{
	return(false);
	}
}

function DisplayFilter()
{
var blnShow = (document.all("divFilter").innerText == "Show Filter");

  if (blnShow) {
     document.all("tblFilter").style.display = "";
     document.all("divFilter").innerText = "Hide Filter";
     document.all("divFilter").title = "Hide Filter";
  } else {
     document.all("tblFilter").style.display = "none";
     document.all("divFilter").innerText = "Show Filter";
     document.all("divFilter").title = "Show Filter";
  }
  document.frmData.blnShowFilter.value = blnShow;
	
return(false);
}

function DisplaySummary()
{
var blnShow = (document.all("divSummary").innerText == "Show Summary");

  if (blnShow) {
     document.all("tblSummary").style.display = "";
     document.all("divSummary").innerText = "Hide Summary";
     document.all("divSummary").title = "Hide Summary";
  } else {
     document.all("tblSummary").style.display = "none";
     document.all("divSummary").innerText = "Show Summary";
     document.all("divSummary").title = "Show Summary";
  }
  document.frmData.blnShowSummary.value = blnShow;

return(false);
}

function DisplayTitle(theLink)
{
	window.status = theLink.title;
	return true;
}

function ClearTitle()
{
	window.status='';
	return true;
}

var strClass;
var strColor;

function HighlightColor(theItem)
{
	strColor = theItem.style.color;
	theItem.style.color = 'Yellow';
}

function deHighlightColor(theItem)
{
	theItem.style.color = strColor;
}


function doMouseOverRow(theRow) {
	strClass = theRow.className;
	theRow.className = "tdHighlight";
    }

function doMouseOutRow(theRow) {
	theRow.className = strClass;
    }
    
function doMouseClickRow(theURL) {
	document.URL = theURL;
    }
