<% Call DebugRecordTime("Starting footer . . .") %>
      <table width="952" border="0" cellspacing="0" cellpadding="3">
		<tr>
		<td class="tdTopBanner" valign="bottom" align="center" style="width: 140px">
		&nbsp;</td>
		<td class="tdTopBanner" valign="bottom" align="center">
		<a class="footer" href="<%= C_HomePath %>">Home</a>
		 | <a class="footer" href="<%= C_HomePath %>search.asp">Search</a>
		 | <a class="footer" href="<%= C_HomePath %>order.asp">Current Order</a>
		<% If iSaveCartActive = 1 Then %> | <a href="<%= C_HomePath %>savecart.asp"><%=Application("CartName")%></a><%End If%>
		 | <a class="footer" href="../../include_commonElements/sitemap.asp">Site Map</a>
		 | <a class="footer" href="../../include_commonElements/privacy.asp">Privacy</a>
		</td>
		</tr>
		<tr>
		<td class="tdTopBanner" align="center">
	&nbsp;</td>
		<td class="tdTopBanner" align="center">
	<a class="footer" href="<%= C_HomePath %>" onclick="window.external.AddFavorite('http://www.sandshot.net/','Sandshot Software - Largest Selection of add-ons for StoreFront by LaGarde'); return false;" title="Bookmark this site">Bookmark Us</a>
		</td>
		</tr>
      </table>
<% If Len(cstrGoogleAnalytics_uacct) > 0 Then %>
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "<%= cstrGoogleAnalytics_uacct %>";
urchinTracker();
</script>
<% End If %>
<script language="javascript" type="text/javascript">setJSErrorHandling();</script>
<% Call DebugRecordTime("Footer complete.") %>