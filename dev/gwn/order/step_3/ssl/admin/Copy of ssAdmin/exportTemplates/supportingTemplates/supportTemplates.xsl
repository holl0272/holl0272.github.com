<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<!--
This file contains supporting Common variables and templates

Common variables:

	outputDelimeter_CR: carriage return
	outputDelimeter_TAB: tab
	
Common templates:

	escape-CSV: escapes commas with quotes
	escape-quote: escapes quotes with double quotes

	stripchars: removes all characters except specified valid characters
		Parameters: x, y, validCharacters

	stripComma: removes commas from the input string

	replace-string:
		Parameters: text, replace, with

-->

<xsl:key name="keyProductsInOrder" match="/orders/Products/Product" use="@uid"/>

<!-- **** Common variables *************************************************************** -->
<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>
<xsl:variable name="outputDelimeter_TAB"><xsl:text>&#9;</xsl:text></xsl:variable>
<xsl:variable name="outputDelimeter_Quote"><xsl:text>"</xsl:text></xsl:variable>

<!-- **** customNumber *************************************************************** -->
<xsl:variable name="currencySymbol" select="orders/order/currencySymbol" />
<xsl:variable name="decimalSeparator" select="orders/order/decimalSeparator" />

<xsl:decimal-format name="customNumber_StandardFormat" decimal-separator="." grouping-separator="," />
<xsl:decimal-format name="customNumber_decimalSeparator" decimal-separator="," grouping-separator="." />

<xsl:template name="customCurrency">
	<xsl:param name='currentNode' />
	<xsl:value-of select="$currencySymbol"/>
	<xsl:call-template name="customNumber">
		<xsl:with-param name="currentNode" select="$currentNode" />
	</xsl:call-template>
</xsl:template>

<xsl:template name="customNumber">
	<xsl:param name='currentNode' />
	<xsl:if test="$decimalSeparator!='true'"><xsl:value-of select="format-number(number($currentNode), '#0.00', 'customNumber_StandardFormat')"/></xsl:if>
	<xsl:if test="$decimalSeparator='true'"><xsl:value-of select="format-number(number($currentNode), '#0,00', 'customNumber_decimalSeparator')"/></xsl:if>
</xsl:template>

<!-- **** csv-entry *************************************************************** -->

	<xsl:template match="node()" mode="csv-entry" name="csv-entry">
	<xsl:param name="string" select="." />

		<xsl:text>"</xsl:text>
		<xsl:choose>
			<xsl:when test="contains($string, '&quot;')">
				<xsl:value-of select="substring-before($string, '&quot;')" />
				<xsl:text>""</xsl:text>
				<xsl:call-template name="escape-CSV">
					<xsl:with-param name="string"
						select="substring-after($string, '&quot;')" />
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$string" />
			</xsl:otherwise>
		</xsl:choose>
		<xsl:text>"</xsl:text>
	</xsl:template>

<!-- **** escape-quote *************************************************************** -->

	<xsl:template match="node()" mode="escape-quote" name="escape-quote">
	<xsl:param name="string" select="." />
	<xsl:choose>
		<xsl:when test="contains($string, '&quot;')">
			<xsl:value-of select="substring-before($string, '&quot;')" />
			<xsl:text>""</xsl:text>
			<xsl:call-template name="escape-CSV">
				<xsl:with-param name="string"
					select="substring-after($string, '&quot;')" />
			</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$string" />
		</xsl:otherwise>
	</xsl:choose>
	</xsl:template>

<!-- **** escape-CSV *************************************************************** -->

	<xsl:template match="node()" mode="escape-CSV" name="escape-CSV">
	<xsl:param name="string" select="." />
	<xsl:choose>
		<xsl:when test="contains($string, '&quot;')">
			<xsl:value-of select="substring-before($string, '&quot;')" />
			<xsl:text>""</xsl:text>
			<xsl:call-template name="escape-CSV">
				<xsl:with-param name="string"
					select="substring-after($string, '&quot;')" />
			</xsl:call-template>
		</xsl:when>
		<xsl:when test="contains($string, ',')">
			<xsl:text>"</xsl:text>
			<xsl:value-of select="substring-before($string, ',')" />
			<xsl:text>,</xsl:text>
			<xsl:call-template name="escape-CSV">
				<xsl:with-param name="string"
					select="substring-after($string, ',')" />
			</xsl:call-template>
			<xsl:text>"</xsl:text>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$string" />
		</xsl:otherwise>
	</xsl:choose>
	</xsl:template>

<!-- **** Product Element *************************************************************** -->
<xsl:template name="productElement">
	<xsl:param name='productUID' />
	<xsl:param name='nodeName' />

	<xsl:for-each select="key('keyProductsInOrder', $productUID)">
		<xsl:for-each select="//*[name() = $nodeName]">
			<!-- The if test is because it will match all nodes otherwise -->
			<xsl:if test="../uid = $productUID"><xsl:value-of select="." /></xsl:if>
		</xsl:for-each>	
	</xsl:for-each>	
</xsl:template>

<!-- **** string-pad-left *************************************************************** -->

	<xsl:template name="string-pad-left">
		<xsl:param name="text"/>
		<xsl:param name="length"/>
		<xsl:param name="padChar"/>
		
		<xsl:if test="number(string-length($text)) &lt; $length">
			<xsl:call-template name="for.loop.string-pad-left">
				<xsl:with-param name="text" select="$text" />
				<xsl:with-param name="padChar" select="$padChar" />
				<xsl:with-param name="i"><xsl:value-of select="number(string-length($text))+1" /></xsl:with-param>
				<xsl:with-param name="count"><xsl:value-of select="number($length)-1" /></xsl:with-param>
			</xsl:call-template>
		</xsl:if>
		<xsl:value-of select="$text"/>
	</xsl:template>

	<!-- this loop is to print out the code/qty for each of the numbered items -->
	<xsl:template name="for.loop.string-pad-left">
		<xsl:param name="padChar" />
		<xsl:param name="i" />
		<xsl:param name="count" />

		<xsl:value-of select="$padChar" />
		
		<xsl:if test="$i &lt;= $count">
			<xsl:call-template name="for.loop.string-pad-left">
				<xsl:with-param name='padChar' select='$padChar' />
				<xsl:with-param name="i"><xsl:value-of select="$i + 1"/></xsl:with-param>
				<xsl:with-param name="count"><xsl:value-of select="$count"/></xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

<!-- **** stripchars *************************************************************** -->

	<xsl:template name="stripchars">
		<xsl:param name="x" />
		<xsl:param name="y" />
		<xsl:param name="validCharacters" />
		    
		<xsl:if test="contains($validCharacters, $x)">
			<xsl:value-of select="$x" />
		</xsl:if>
		    
		<xsl:if test="string-length($y) > 0">
			<xsl:call-template name="stripchars">
			<xsl:with-param name="x" select="substring($y, 1, 1)" /> 
			<xsl:with-param name="y" select="substring($y, 2, string-length($y))" /> 
			<xsl:with-param name="validCharacters" select="$validCharacters" /> 
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

<!-- **** stripComma *************************************************************** -->

	<xsl:template match="node()" mode="stripComma" name="stripComma">
		<xsl:call-template name="replace-string">
			<xsl:with-param name="text" select="." />
			<xsl:with-param name="replace" select="','" />
			<xsl:with-param name="with" select="''" />
		</xsl:call-template>
	</xsl:template>

<!-- **** stripQuote *************************************************************** -->

	<xsl:template match="node()" mode="stripQuote" name="stripQuote">
		<xsl:call-template name="replace-string">
			<xsl:with-param name="text" select="." />
			<xsl:with-param name="replace" select="'&quot;'" />
			<xsl:with-param name="with" select="''" />
		</xsl:call-template>
	</xsl:template>

<!-- **** replace-string *************************************************************** -->

	<xsl:template name="replace-string">
		<xsl:param name="text"/>
		<xsl:param name="replace"/>
		<xsl:param name="with"/>
		<xsl:choose>
			<xsl:when test="contains($text,$replace)">
				<xsl:value-of select="substring-before($text,$replace)"/>
				<xsl:value-of select="$with"/>
				<xsl:call-template name="replace-string">
				<xsl:with-param name="text" select="substring-after($text,$replace)"/>
				<xsl:with-param name="replace" select="$replace"/>
				<xsl:with-param name="with" select="$with"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$text"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>