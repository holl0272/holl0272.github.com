
<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% Call DebugRecordTime("Starting left nav . . .") %>
	<% 'Call DisplayMiniCart %>

      <table class="clsPart" cellpadding="0" cellspacing="0" width="150" border="0" ID="tblCategoryMenu">
		<tr>
		  <td align="center" class="lclsMenuBody">
			<table class="lclsMenuBody" width="100%" cellpadding="0" cellspacing="0" border="0" ID="tblCategoryMenu_Inner">
			  <div style="height:20px">&nbsp;</div>
			  <% Call WriteCategories(mstrCurrentCategory, mstrCurrentSubCategory, mstrCurrentSubSubCategory) %>
			  <tr>
			    <td valign="middle" class="lclsMenuNavigationCat" style="height: 30px">
			    &nbsp;</td>
			  </tr>
			  <tr>
			  <td valign="middle" class="lclsMenuNavicationCat">
			  	<a class="lclsMenuNavigationCat" href="Shipping.asp">When will my jerseys arrive?</a>
			  	</td>
			  	</tr>
			  <tr>
			  <td valign="middle" class="lclsMenuNavicationCat">
			  	<a class="lclsMenuNavigationCat" href="Lettering.asp">Custom lettering options</a>
			  	</td>
			  	</tr>
			  	<tr>
			  <td valign="middle" class="lclsMenuNavicationCat">
			  	<a class="lclsMenuNavigationCat" href="sizing.asp">Sizing information</a>
			  	</td>
			  	</tr>
				<tr>
			    <td valign="middle" class="lclsMenuNavigationCat" style="height: 40px">
			    &nbsp;</td>
			  </tr>

			<tr>
			    <td valign="middle" class="lclsMenuNavigationCat">
			  	  <a class="lclsMenuNavigationCat" href="MailSubscribe.asp">Newsletter</a>
			    </td>
			  </tr>
			  <tr>
			    <td valign="middle" class="lclsMenuNavigationCat">
			  	  <a class="lclsMenuNavigationCat" href="myAccount.asp">My Account</a>
			    </td>
			  </tr>
			  <tr>
			    <td valign="middle" class="lclsMenuNavigationCat">
			  	  <form id='frmCategory' name='frmCategory' action='../search_results.asp' method='Get' onsubmit="" style="display:inline"><input type="text" name="txtsearchParamTxt" ID="txtsearchParamTxt_frmCategory" value="SEARCH" onfocus="this.value='';" size="8" />
										<a href="<%= C_HomePath %>advancedSearch.asp" onclick="if (document.frmCategory.txtsearchParamTxt.value=='SEARCH'){alert('Please enter a word to search for.'); document.frmCategory.txtsearchParamTxt.focus(); return false;} document.frmCategory.submit(); return false;">
                                        <img src='images/buttons/go3.gif' alt='Search' border='0'></a><br />
                                        <font size="-2"><a href="advancedSearch.asp" title="Search using advanced options">Advanced Search</a></font>
										<input type="hidden" name="txtFromSearch" id="txtFromSearch_frmCategory" value="fromSearch">
										<input type="hidden" name="txtsearchParamType" id="txtsearchParamType_frmCategory" value="ALL">
										<input type="hidden" name="txtsearchParamMan" id="txtsearchParamMan_frmCategory" value="ALL">
										<input type="hidden" name="txtsearchParamVen" id="txtsearchParamVen_frmCategory" value="ALL">
										<input type="hidden" name="txtsearchParamCat" id="txtsearchParamCat_frmCategory" value="ALL">
										<input type="hidden" name="subcat" id="subcat_frmCategory" value="">
										<input type="hidden" name="iLevel" id="iLevel_frmCategory" value="1">
										<input type="hidden" name="txtCatName" id="txtCatName_frmCategory" value="">
										</form>
			    </td>
			  </tr>
			  <tr>
			    <td valign="middle" class="lclsMenuNavigationCat" height="30">
			  	  &nbsp;</td>
			  </tr>
			  <tr>
			    <td valign="middle" class="style1" height="60">
			  	  GameWearNow &amp; <br>
					<img alt="Sports Spot" src="../images/ss%20logo%20IIsmall.gif" width="138" height="18"><br>
					Custom jerseys since 1955</td>
			  </tr>
			  <tr>
			    <td valign="middle" class="lclsMenuNavigationCat" height="30">
			  	  &nbsp;</td>
			  </tr>
			  <tr>
			    <td> <p align="center"><script src="https://siteseal.thawte.com/cgi/server/thawte_seal_generator.exe">
    </script></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
<% Call DebugRecordTime("Right left complete.") %>