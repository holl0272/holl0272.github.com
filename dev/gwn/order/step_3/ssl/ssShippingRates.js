function getSelectValue(theSelect)
{

	if (theSelect.selectedIndex == -1)
	{
		return('');
	}else{
		return(theSelect.options[theSelect.selectedIndex].value);
	}

}

function letSelectValue(theSelect,theValue)
{

	for (var i = 0;  i < theSelect.options.length;  i++)
	{
		if (theSelect.options[i].value == theValue)
		{
			theSelect.selectedIndex = i;
			return true;
		}
	}
	
	return false;
}

function ValidateCountry()
{
	var cbytDefaultDomestic = 16;
	var cbytDefaultIntl = 27;		//Global Airmail Parcel Post
	
	var theForm = document.form1;
	var pstrDestCountry;
	var pstrShipMethod = theForm.Shipping.value;

	if (theForm.ShipFirstName.value == "")
	{
		pstrDestCountry = getSelectValue(theForm.Country);
	}
	else
	{
		pstrDestCountry = getSelectValue(theForm.ShipCountry);
	}

	if (pstrDestCountry  == "")	//Exit gracefully so form validation can take over
	{
	return true;
	}
	
	if (pstrShipMethod == "")
	{
	alert("Please select a shipping method using the /View Rates/ button.");
	return false;
	}
	
	if (pstrDestCountry != 'US')
	{
		if (pstrShipMethod == cbytDefaultIntl)	
		{
			return true;
		}
		else
		{
			alert("We use Global Airmail Parcel Post for all non-U.S. orders.");
//			theForm.Shipping.value = cbytDefaultIntl;
//			return true;
			letSelectValue(theForm.Shipping,cbytDefaultIntl);
			theForm.Shipping.focus();
			return false;
		}
	}
	else
	{
		if (pstrShipMethod == cbytDefaultIntl)
		{
			alert("Global Airmail Parcel Post cannot be used for U.S. orders. \n Please reselect your shipping method.");
			letSelectValue(theForm.Shipping,cbytDefaultDomestic);
			theForm.Shipping.focus();
			return false;
		}
		else
		{
			return true;
		}
	}
	
}

function ssGetRates(strDisplayType)
{
var pstrZip;
var pstrDestCountry;
var pstrDestState;
var pstrURL;

var theForm = document.form1;
var p_blnShip = (theForm.ShipFirstName.value != "");

	if (p_blnShip)
	{
		pstrZip = theForm.ShipZip.value;
		pstrDestCountry = getSelectValue(theForm.ShipCountry);
		pstrDestState = getSelectValue(theForm.ShipState);
	}else{
		pstrZip = theForm.Zip.value;
		pstrDestCountry = getSelectValue(theForm.Country);
		pstrDestState = getSelectValue(theForm.State);
	}

	if (pstrDestCountry  == "")
	{
		alert("Please select a Country to ship to.");
		if (p_blnShip){theForm.ShipCountry.focus()}else{theForm.Country.focus()}
		return false
	}

	if (((pstrDestCountry == "US") || (pstrDestCountry == "CA")) && (pstrZip == ""))
	{
		alert("Please enter a ZIP code.");
		if (p_blnShip){theForm.ShipZip.focus()}else{theForm.Zip.focus()}
		return false
	}

	var pstrExtraInformation = "";
	//pstrExtraInformation = "&ShipResidential=" + theForm.shipRes[0].checked + "&InsideDelivery=" + theForm.shipIndoorDeliver[0].checked

	pstrURL = "ssShippingRates.asp?DestinationZip=" + pstrZip + "&DestinationCountry=" + pstrDestCountry + "&DestinationState=" + pstrDestState + "&DisplayStyle=" + strDisplayType + pstrExtraInformation;
	PreView=window.open(pstrURL,"Preview","toolbar=0,location=1,directories=0,status=0,menubar=No,copyhistory=0,scrollbars=1,width=400,height=300");
}

function GetRates_SF2k(strDisplayType, subTotal)
{
	var bytTarget = document.forms.length - 1;
	var theForm = document.forms[bytTarget];

	var pstrZip = theForm.DEST_ZIP.value;
	var pstrDestCountry;
	var pstrDestState;
	var pstrURL;

	if (theForm.Country.selectedIndex == -1)
	{
		alert("Please select a Country.");
		theForm.Country.focus()
		return false
	}else{
		pstrDestCountry = theForm.Country[theForm.Country.selectedIndex].value;
	}

	if (theForm.State.selectedIndex != -1)
	{
		pstrDestState = theForm.State[theForm.State.selectedIndex].value;
	}

	if (((pstrDestCountry == "US") || (pstrDestCountry == "CA")) && (pstrZip == ""))
	{
		alert("Please enter a ZIP code.");
		theForm.DEST_ZIP.focus()
		return false
	}

	var pstrExtraInformation;
	//pstrExtraInformation = "&ShipResidential=" + theForm.shipRes[0].checked + "&InsideDelivery=" + theForm.shipIndoorDeliver[0].checked

	pstrURL = "ssShippingRates.asp?DestinationZip=" + pstrZip + "&DestinationCountry=" + pstrDestCountry + "&DestinationState=" + pstrDestState + "&DisplayStyle=" + strDisplayType + "&subTotal=" + subTotal + pstrExtraInformation;
	PreView = window.open(pstrURL,"Preview", "toolbar=No,location=1,directories=No,status=0,menubar=No,copyhistory=0,scrollbars=1,width=400,height=400");
}

