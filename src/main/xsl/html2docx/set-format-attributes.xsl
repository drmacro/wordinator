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
       Mode "set-run-style-attributes"  
       
       Does the mapping from elements-in-context to Word styles.
       
       =================================================================== -->

  <!-- Map of @class names to wp:run format attributes (e.g., bold="true"). -->
  <xsl:variable name="classToRunFormatAttsMap" as="map(xs:string, attribute()*)">
    <xsl:map>
      <xsl:map-entry key="'bold'">
        <xsl:attribute name="bold" select="'true'"/>
      </xsl:map-entry>
      <xsl:map-entry key="'italic'">
        <xsl:attribute name="italic" select="'true'"/>
      </xsl:map-entry>
      <xsl:map-entry key="'title'">
        <xsl:attribute name="italic" select="'true'"/>
      </xsl:map-entry>
      <xsl:map-entry key="'dfn'">
        <xsl:attribute name="italic" select="'true'"/>
      </xsl:map-entry>
    </xsl:map>
  </xsl:variable>
  
  <xsl:template name="set-run-format-attributes" as="attribute()*">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- For text, check the grandparent and then parent element, for 
         elements, check the parent and then the element.
         
         NOTE: The last attribute with a given name takes precendence
      -->
    <xsl:choose>
      <xsl:when test="self::text()">
        <xsl:apply-templates mode="set-run-format-attributes" select="../..">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
        <xsl:apply-templates mode="set-run-format-attributes" select="..">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates mode="set-run-format-attributes" select=".">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
        <xsl:apply-templates mode="set-run-format-attributes" select="..">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" as="attribute()*" 
    match="
    xhtml:dfn | dfn |
    xhtml:i | i
    ">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:attribute name="italic" select="'true'"/>
    
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" as="attribute()*" 
    match="
    xhtml:dt | dt |
    xhtml:b | b
    ">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-run-format-attributes: {name(.)}: Setting bold to true</xsl:message>
    </xsl:if>
    
    <xsl:attribute name="bold" select="'true'"/>
    
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" as="attribute()*" 
    match="
      xhtml:u | u
    ">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-run-format-attributes: {name(.)} "{.}": Setting underline to single</xsl:message>
    </xsl:if>
    
    <xsl:attribute name="underline" select="'single'"/>
    
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" match="xhtml:*[@style] | *[@style]" as="attribute()*" priority="-0.5">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="properties" as="xs:string*" select="tokenize(., ';')"/>
    <xsl:variable name="font-weight" as="xs:string?" select="$properties[starts-with(., 'font-weight:')]"/>
    <xsl:variable name="font-style" as="xs:string?" select="$properties[starts-with(., 'font-style:')]"/>
    
    <xsl:if test="exists($font-weight)">
      <xsl:variable name="value" as="xs:string?" select="local:get-style-value($font-weight)"/>
      <xsl:choose>
        <xsl:when test="$value = ('bold')">
          <xsl:attribute name="bold" select="'true'"/>    
        </xsl:when>
        <xsl:when test="$value = ('normal')">
          <xsl:attribute name="bold" select="'false'"/>    
        </xsl:when>
      </xsl:choose>      
      
    </xsl:if>
    <xsl:if test="exists($font-style)">
      <xsl:variable name="value" as="xs:string?" select="local:get-style-value($font-style)"/>
      <xsl:choose>
        <xsl:when test="$value = ('italic')">
          <xsl:attribute name="italic" select="'true'"/>    
        </xsl:when>
        <xsl:when test="$value = ('normal')">
          <xsl:attribute name="italic" select="'false'"/>    
        </xsl:when>
      </xsl:choose>      
      
    </xsl:if>
    
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" match="xhtml:*[@class] | *[@class]" as="attribute()*" priority="-0.8">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-run-format-attributes *[@class]: @class="{@class}"</xsl:message>
    </xsl:if>
    <xsl:variable name="classTokens" as="xs:string*" select="tokenize(@class, ' ')"/>
    <xsl:variable name="atts" as="attribute()*" 
      select="$classTokens ! map:get($classToRunFormatAttsMap, .)"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-run-format-attributes *[@class]: result atts:
<xsl:sequence select="$atts"/>        
      </xsl:message>
    </xsl:if>
    <xsl:sequence select="$atts"/>
    
  </xsl:template>
  
  <xsl:template mode="set-run-format-attributes" match="* | text() | @*" as="attribute()*" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <!-- Nothing to do -->
  </xsl:template>
  
</xsl:stylesheet>