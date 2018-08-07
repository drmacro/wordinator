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
       Mode "get-style-name
       
       Does the mapping from elements-in-context to Word styles.
       
       =================================================================== -->
  
  
  <xsl:variable name="classToStyleNameMap" as="map(xs:string, xs:string)">
    <xsl:map>
      <xsl:map-entry key="'p1'" select="'Paragraph 1'"/>
    </xsl:map>
  </xsl:variable>
    
  <xsl:template mode="get-style-name" match="xhtml:section/xhtml:header" as="xs:string?">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <xsl:variable name="headingLevel" as="xs:integer"
      select="count(ancestor::xhtml:section)"
    />
    <xsl:variable name="result" as="xs:string" 
      select="'Heading ' || $headingLevel"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] get-style-name: section/header. Returning "{$result}".</xsl:message>
    </xsl:if>
    <xsl:sequence select="$result"/>
  </xsl:template>
  
  <xsl:template mode="get-style-name" as="xs:string?"
    match="
      xhtml:h1 |
      xhtml:h2 |
      xhtml:h3 |
      xhtml:h4 |
      xhtml:h5
    " 
    >
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="heading-number" as="xs:string"
      select="substring-after(name(.), 'h')"
    />
    <xsl:variable name="headingLevel" as="xs:integer"
      select="xs:integer($heading-number)"
    />
    <xsl:variable name="result" as="xs:string" 
      select="'Heading ' || $headingLevel"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] get-style-name: {name(.)}. Returning "{$result}".</xsl:message>
    </xsl:if>
    <xsl:sequence select="$result"/>
  </xsl:template>
  
  <xsl:template mode="get-style-name" match="xhtml:span[@class] | xhtml:p[@class]" as="xs:string?">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="tokens" as="xs:string*" select="tokenize(@class, ' ')"/>
    <xsl:variable name="key" select="$tokens[1]"/>
    <xsl:variable name="styleName" as="xs:string?"
      select="map:get($classToStyleNameMap, $key)"
    />
    <xsl:sequence select="if (exists($styleName)) then $styleName else ()"/>
  </xsl:template>
  
  <xsl:template mode="get-style-name" match="xhtml:li" as="xs:string">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="list-type" as="xs:string" select="name(..)"/>
    <!-- FIXME: This is not good enough but will work for the code index, which only uses ul -->
    <xsl:variable name="level" as="xs:string"
      select="string(count(ancestor::*[name(.) = $list-type]))"
    />
    <xsl:variable name="level" as="xs:string"
      select="if ($level eq '1') then '' else $level"
    />
    <xsl:sequence select="'List ' || $level"/>
  </xsl:template>
  
  <xsl:template mode="get-style-name" match="*" priority="-1" as="xs:string?">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- No style mapping -->
  </xsl:template>
</xsl:stylesheet>