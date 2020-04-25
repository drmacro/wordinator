<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  exclude-result-prefixes="xs"
  version="3.0">
  
  <!-- Test use of catalogs with Saxon to resolve URIs -->
  
  <xsl:import href="urn:xslt:wordinator:html2docx.xsl"/>
  
  <xsl:template match="/">
    <xsl:message>+ [DEBUG] "/" test-catalog-resolution.xsl</xsl:message>
    <xsl:next-match/>
  </xsl:template>
</xsl:stylesheet>