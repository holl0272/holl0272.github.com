<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match='/'>
    <xsl:apply-templates select="orders/order" />
  </xsl:template>

  <xsl:template match="order" >

Thank you for ordering from Space Universe!  

We are busy processing your order,  most orders are shipped within 24 hrs.  

If you have any questions you can contact us Monday through Friday, 

8am to 5pm PST by calling 888-73-SPACE

-----------------
Sold To
-----------------
<xsl:apply-templates select="billingAddress" />

<xsl:text>&#10;</xsl:text>
<xsl:value-of select="orderPaymentMethod"/>

-----------------
Shipped To
-----------------
<xsl:apply-templates select="shippingAddress" />

-----------------
Purchase Summary
-----------------
Order ID: <xsl:value-of select="orderID"/>
<xsl:text>&#10;</xsl:text>

<xsl:for-each select="orderDetail">
Item <xsl:number value="position()" format="1"/>
Product ID: <xsl:value-of select="odrdtProductID" />
Product Name: <xsl:value-of select="odrdtProductName" />
	<xsl:for-each select="odrdtAttDetailID">
		<xsl:text>  </xsl:text><xsl:value-of select="odrattrName" />: <xsl:value-of select="odrattrAttribute" />
	</xsl:for-each>
Product Price: $<xsl:value-of select="odrdtSubTotal" />
Quantity: <xsl:value-of select="odrdtQuantity" />
<xsl:text>&#10;</xsl:text>
</xsl:for-each>
SubTotal:    $<xsl:value-of select="orderAmount" />
Shipping:    $<xsl:value-of select="orderShippingAmount" /> (<xsl:value-of select="orderShipMethod" />)
Handling:    $<xsl:value-of select="orderHandling" />
State Tax:   $<xsl:value-of select="orderSTax" />
Country Tax: $<xsl:value-of select="orderCTax" />
Grand Total: $<xsl:value-of select="orderGrandTotal" />

NOTICE: The shipping costs as shown above do not necessarily represent the carrier's published rates and may include additional charges levied by the merchant.

Special Instructions: <xsl:value-of select="orderComments" />
User Name: <xsl:value-of select="billingAddress/Email" />
Password: <xsl:value-of select="billingAddress/custPasswd" />
Use this information next time you order for quick access to your Customer Information.

  </xsl:template>

  <xsl:template match="shippingAddress" >
	<xsl:value-of select="FirstName"/><xsl:text> </xsl:text>
	<xsl:choose>
		<xsl:when test="MiddleInitial[.&gt;'']">
			<xsl:value-of select="MiddleInitial" />
		</xsl:when>
	</xsl:choose>
	<xsl:value-of select="LastName" />
	<xsl:text>&#10;</xsl:text>
	
	<xsl:choose>
		<xsl:when test="Company[.!='']">
			<xsl:value-of select="Company" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Addr1[.!='']">
			<xsl:value-of select="Addr1" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Addr2[.!='']">
			<xsl:value-of select="Addr2" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:value-of select="City"/>,<xsl:text> </xsl:text><xsl:value-of select="State"/><xsl:text> </xsl:text><xsl:value-of select="Zip"/>

	<xsl:choose>
		<xsl:when test="Country[.='US']">
			<xsl:text>&#10;</xsl:text>
			<xsl:value-of select="CountryName" />
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Phone[.!='']">
			Phone: <xsl:value-of select="Phone" />
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Fax[.!='']">
			Fax: <xsl:value-of select="Fax" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Email[.!='']">
			<xsl:value-of select="Email" />
		</xsl:when>
	</xsl:choose>

  </xsl:template>

  <xsl:template match="billingAddress" >
	<xsl:value-of select="FirstName"/><xsl:text> </xsl:text>
	<xsl:choose>
		<xsl:when test="MiddleInitial[.&gt;'']">
			<xsl:value-of select="MiddleInitial" />
		</xsl:when>
	</xsl:choose>
	<xsl:value-of select="LastName" />
	<xsl:text>&#10;</xsl:text>
	
	<xsl:choose>
		<xsl:when test="Company[.!='']">
			<xsl:value-of select="Company" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Addr1[.!='']">
			<xsl:value-of select="Addr1" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Addr2[.!='']">
			<xsl:value-of select="Addr2" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:value-of select="City"/>,<xsl:text> </xsl:text><xsl:value-of select="State"/><xsl:text> </xsl:text><xsl:value-of select="Zip"/>

	<xsl:choose>
		<xsl:when test="Country[.='US']">
			<xsl:text>&#10;</xsl:text>
			<xsl:value-of select="CountryName" />
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Phone[.!='']">
			Phone: <xsl:value-of select="Phone" />
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Fax[.!='']">
			Fax: <xsl:value-of select="Fax" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
	
	<xsl:choose>
		<xsl:when test="Email[.!='']">
			<xsl:value-of select="Email" />
			<xsl:text>&#10;</xsl:text>
		</xsl:when>
	</xsl:choose>
  </xsl:template>

</xsl:stylesheet>