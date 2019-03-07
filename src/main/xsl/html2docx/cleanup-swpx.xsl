<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:wp="urn:ns:wordinator:simplewpml"
  xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:local="urn:ns:local-functions"
  xmlns:relpath="http://dita2indesign/functions/relpath"
  xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  exclude-result-prefixes="xs local relpath xhtml map"
  version="3.0"
  expand-text="yes"
  >
  <!-- ===================================================================
       Mode "cleanup-swpx"
       
       Cleans up the generated SWPX to try to make sure it's valid.
       
       =================================================================== -->
  
  <xsl:template mode="cleanup-swpx" match="wp:run[empty(node())]" priority="10">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:message>- [WARN] Empty run: <xsl:sequence select="."/> (preceding: <xsl:sequence select="./preceding-sibling::*[1]"/>)</xsl:message>
    <!-- Eliminate completely empty runs -->
  </xsl:template>


  <xsl:template mode="cleanup-swpx" match="wp:body/wp:run">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:message>- [WARN] Run found in body: <xsl:sequence select="."/></xsl:message>
    <wp:p>
      <xsl:copy>
        <xsl:apply-templates mode="#current" select="@*, node()">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
      </xsl:copy>
    </wp:p>
  </xsl:template>
  
  <xsl:template mode="cleanup-swpx" match="wp:p/wp:p">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:message>- [WARN] wp:p found in wp:p: <xsl:sequence select="."/></xsl:message>
    
    <xsl:apply-templates mode="#current" select="node()">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="text() | @* | processing-instruction()" mode="cleanup-swpx" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:sequence select="."/>
  </xsl:template>
  
  <xsl:template match="*" mode="cleanup-swpx" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:copy>
      <xsl:apply-templates mode="#current" select="@*, node()">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>