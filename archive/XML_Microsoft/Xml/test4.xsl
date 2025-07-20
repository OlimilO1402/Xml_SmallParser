<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
           version="1.0">

  <!-- Map "root" element to "network" element. -->
  <xsl:template match="root">
    <xsl:element name="network">
       <xsl:apply-templates/>
    </xsl:element>   
  </xsl:template>

  <!-- Keep any other elements as-is. -->
  <xsl:template match="/ | @* | node()">
    <xsl:copy>
      <xsl:apply-templates select="@* | node()"/>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>