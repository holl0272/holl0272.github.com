var maryCurrentAttributePrice = new Array();
var maryAttributePrice = new Array();
var prodBasePrice = 0;
var cstrBaseImagePath = "";

function updateProductPrice(theSelect, attrNum)
{
	var theForm = theSelect.form;
	var attrIndex = theSelect.selectedIndex + 1;
	var dblPrice;

	letCurrentAttributePrice(attrNum, attrIndex);
	dblPrice = getProductPrice(attrNum, attrIndex);
	theForm.prodPrice.value = dblPrice;
	//alert(dblPrice);
}

function setProductPrice(theSelect, attrNum)
{
	var theForm = theSelect.form;
	var attrIndex = theSelect.selectedIndex + 1;
	var dblPrice;

	letCurrentAttributePrice(attrNum, attrIndex);
	dblPrice = getProductPrice(attrNum, attrIndex);
	theForm.prodPrice.value = dblPrice;
	//alert(dblPrice);
}

function setCurrentAttributePrice(attrNum, attrPrice)
{
	maryCurrentAttributePrice[attrNum] = attrPrice;
}

function setAttributePrice(attrNum, attrIndex, attrPrice)
{
	if (attrIndex == 1){ maryAttributePrice[attrNum] = new Array(); }
	maryAttributePrice[attrNum][attrIndex] = attrPrice;
}

function letCurrentAttributePrice(attrNum, attrIndex)
{
	maryCurrentAttributePrice[attrNum] = maryAttributePrice[attrNum][attrIndex];
}

function getProductPrice()
{
var pdblPrice;

	pdblPrice = prodBasePrice;
	for (var i = 1; i < maryCurrentAttributePrice.length; i++)
	{
		pdblPrice = pdblPrice + maryCurrentAttributePrice[i];
	}
	pdblPrice = formatDollar(pdblPrice,"");

	return pdblPrice;
}

function formatDollar (Val, DollarSign)
{

	Val= "" + Val;

	if (Val.indexOf (".", 0)!=-1)
	{
		Dollars = Val.substring(0, Val.indexOf (".", 0));
    	Cents = Val.substring(Val.indexOf (".", 0)+1, Val.indexOf (".", 0)+3);
    	if (Cents.length==0) 			Cents="00";
    	if (Cents.length==1)			Cents=Cents+"0";
	}else{
    	Dollars = Val;
    	Cents = "00";
	}

    if (DollarSign)
    {
    	OutString = ("$"+Dollars+"."+Cents);
    }else{
    	OutString = (Dollars+"."+Cents);
    }

	len=Dollars.length;
	if (len>=3)
	{
		OutString="";
		while (len>0)
		{
			TempString=Dollars.substring(len-3, len)
    		if (TempString.length==3)
    		{
    			OutString = "," + TempString + OutString;
    			len=len-3;
    		}else{
    			OutString=TempString+OutString;
    			len=0;
    		}
   		}

    	if (OutString.substring(0, 1)==",")
    	{
    		Dollars=OutString.substring (1, OutString.length);
    	}else{
    		Dollars=OutString
    	}

    	if (DollarSign)
    	{
    		OutString = ("$"+Dollars+"."+Cents);
    	}else{
    		OutString = (Dollars+"."+Cents);
    	}
    }

    return OutString;
}

var theCustomImage = new Array();	//this is set by the image
var maryAttrImages = new Array();
var maryCustomImages = new Array();

function setCustomImage(theImage, strProdID){theCustomImage[strProdID] = theImage;}

function changeCustomImage(strProdID, theSelect){theCustomImage[strProdID].src = maryAttrImages[strProdID][theSelect[theSelect.selectedIndex].value];}

function letCustomImageEntry(lngIndex, strImage){maryAttrImages[lngIndex] = strImage;}

function letCustomImage(lngIndex, strImage)
{
	letCustomImageEntry(lngIndex, strImage);
	theCustomImage.src = getCustomImagePath();
}

function changeDetailImage(strProdID, strPath){theCustomImage[strProdID].src = strPath;}

function getCustomImagePath()
{
var pstrBaseImagePath;

	for (var i = 0; i < maryAttrImages.length; i++)
	{
		if (maryAttrImages[i] != undefined)
		{
			if (maryAttrImages[i].length > 0)
			{
				pstrBaseImagePath = "_" + maryAttrImages[i];
			}
		}
	}

	if (pstrBaseImagePath != ""){ pstrBaseImagePath = cstrBaseImagePath + pstrBaseImagePath + ".jpg"; }
	//alert(pstrBaseImagePath);
	return pstrBaseImagePath;
}

function setAttributeOptional(theFormName, theFieldName, optional)
{
	var e = document.getElementById(theFieldName);
	e.optional = optional;
	//alert(e.name + ": " + e.optional);
}

function setAttributeTitle(theFormName, theFieldName, theTitle)
{
	//var e = eval(theFormName + '.' + theFieldName);
	var e = document.getElementById(theFieldName);
	e.title = theTitle;
	//alert(e.title);
}

function show_BigImage(path){
	var sFeatures, h, w, win, i
	h = window.screen.availHeight
	w = window.screen.availWidth
	sFeatures = "height=" + h*.50 + ",width=" + w*.52 + ",screenY=" + (h*.30) + ",screenX=" + (w*.33) + ",top=" + (h*.30) + ",left=" + (w*.33) + ",resizable=yes"
	win = window.open(path,"BigImage",sFeatures)
}

function validateQty(theQuantityBox, minQty)
{
	var checkValue;

	if (minQty.length == 0)
	{
		checkValue = 1;
	}else{
		checkValue = minQty;
	}

	if (!isInteger(theQuantityBox, true, 'Please enter an integer greater than one for the quantity'))
	{
		return false;
	}

	if (theQuantityBox.value < checkValue)
	{
		alert("The minimum quantity for this product is " + checkValue + ".");
		theQuantityBox.value = checkValue;
		theQuantityBox.select();
		return false;
	}

	return true;
}

function hideElement(theItem){theItem.style.display="none";}
function showElement(theItem){theItem.style.display="";}
function showHideElement(theItem)
{
	if (theItem.style.display=="")
	{
		hideElement(theItem);
	}else{
		showElement(theItem);
	}
}

/*
	mfsct - Microformats scanner class tool
	written by Chris Heilmann (http://icant.co.uk) building on scripts by
	Jonathan Snook, http://www.snook.ca/jonathan and
	Robert Nyman, http://www.robertnyman.com
*/
mfsct = {
	check:function(oElm, strClassName){
	    strClassName = strClassName.replace(/\-/g, "\\-");
	    var oRegExp = new RegExp("(^|\\s)" + strClassName + "(\\s|$)");
		return oRegExp.test(oElm.className);
	},
	add:function(oElm, strClassName){
		if(!mfsct.check(oElm, strClassName)){
			oElm.className+=oElm.className?' '+strClassName:strClassName;
		}
	},
	remove:function(oElm, strClassName){
		var rep=oElm.className.match(' '+strClassName)?' '+strClassName:strClassName;
	    oElm.className=oElm.className.replace(rep,'');
	    oElm.className.replace(/^\s./,'');
	},
	display:function(o){
		if(o.className){o = o.className;}
		if(window.console){
			window.console.log(o);
		} else {
			alert(o);
		}
	},
	getElements:function(oElm, strTagName, strClassName){
	    var arrElements = (strTagName == "*" && oElm.all)? oElm.all : oElm.getElementsByTagName(strTagName);
	    var arrReturnElements = [];
	    for(var i=0; i<arrElements.length; i++){
	        var temp = arrElements[i];
			if(mfsct.check(temp, strClassName)){
				arrReturnElements.push(temp);
			}
	    }
	    return (arrReturnElements)
	}
}