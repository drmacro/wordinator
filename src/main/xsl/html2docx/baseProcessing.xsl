<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:wp="urn:ns:wordinator:simplewpml"
  xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:local="urn:ns:local-functions"
  xmlns:relpath="http://dita2indesign/functions/relpath"
  exclude-result-prefixes="xs local relpath xhtml"
  version="3.0"
  expand-text="yes"
  >
  <!-- ===================================================================
       Base processing for XHTML to DOCX
       
       =================================================================== -->
  
  <xsl:key name="elementsById" match="*[@id]" use="@id"/>
  
 
  <xsl:template match="xhtml:html | html">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
<!--    <xsl:variable name="doDebug" as="xs:boolean" select="true()"/>-->
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] #default: (not a chunk) Handling {name(..)}/{name(.)}</xsl:message>
    </xsl:if>
    
    <xsl:apply-templates>
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="xhtml:head | head | body/header |
                       xhtml:body/xhtml:header"
    >
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <!-- No direct output from head -->
  </xsl:template>
  
  <xsl:template 
    match="xhtml:section[local:is-chunk(.)] | section[local:is-chunk(.)] | 
           xhtml:html[local:is-chunk(.)] | html[local:is-chunk(.)]" priority="10">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
        
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] #default: is-chunk() is true: Handling {name(..)}/{name(.)}</xsl:message>
    </xsl:if>
        
    <xsl:variable name="result-uri" as="xs:string"
      select="local:get-result-uri(., $chunk-level, $outputDirectory)"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [INFO] Generating result document "{$result-uri}"...</xsl:message>
    </xsl:if>
    <xsl:variable name="swpx-base-result" as="element()">
      <wp:document>
        <wp:page-sequence-properties>
          <!-- Use default page numbering properties -->
          <xsl:if test="$doDebug">
            <xsl:message>+ [DEBUG] page-sequence properties: applying templates to {name(.)} in mode make-section-header-and-footer...</xsl:message>
          </xsl:if>
          <xsl:apply-templates mode="make-section-header-and-footer" select=".">
            <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
          </xsl:apply-templates>
          <xsl:if test="$doDebug">
            <xsl:message>+ [DEBUG] applying templates to {name(.)} Done with page-sequence-properties.</xsl:message>
          </xsl:if>
        </wp:page-sequence-properties>          
        <wp:body>
          <xsl:apply-templates select="." mode="make-body-content">
            <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
          </xsl:apply-templates>
        </wp:body>
      </wp:document>           
    </xsl:variable>
    <xsl:if test="$doDebug or false()">
      <debug>
        <message>swpx-base-result for {$result-uri}</message>
        <xsl:sequence select="$swpx-base-result"/>
      </debug>
    </xsl:if>
    <xsl:result-document href="{$result-uri}" format="swpx" >
      <xsl:apply-templates select="$swpx-base-result" mode="cleanup-swpx">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </xsl:result-document>
  </xsl:template>
  
  <xsl:template mode="make-body-content" match="xhtml:html | html">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] make-body-content: Handling {name(..)}/{name(.)}</xsl:message>
    </xsl:if>
    
    <xsl:apply-templates mode="#default" select="xhtml:body | body">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template mode="make-body-content" match="xhtml:section | section">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] make-body-content: Handling {name(..)}/{name(.)}</xsl:message>
    </xsl:if>
    
    <xsl:apply-templates mode="#default">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="xhtml:* | *" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug and false()">
      <xsl:message>+ [DEBUG] Fallback handling: {name(..)}/{name(.)}</xsl:message>
    </xsl:if>

    <xsl:apply-templates>
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template match="xhtml:a[empty(@href)] | a[empty(@href)]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:apply-templates>
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template match="xhtml:hr | hr">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <!-- Ignore -->
  </xsl:template>
  
  <xsl:template match="xhtml:a[@href] | a[@href]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="targetId" as="xs:string?" select="tokenize(@href, '#')[last()]"/>
    
    <wp:hyperlink href="#{$targetId}">
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:hyperlink>    
  </xsl:template>
  
  <xsl:template match="xhtml:a[@class = ('footnoteref')] | a[@class = ('footnoteref')]" mode="running-header">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <!-- Ignore -->
  </xsl:template>
  
  <xsl:template match="xhtml:a[@class = ('footnoteref')] | a[@class = ('footnoteref')]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="targetId" as="xs:string?" select="tokenize(@href, '#')[last()]"/>

    <xsl:variable name="aside" as="element(xhtml:aside)?"
      select="key('elementsById', $targetId, root(.))"
    />
    <xsl:apply-templates select="$aside" mode="make-footnote">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
  </xsl:template> 
  
  <xsl:template match="xhtml:aside | aside" mode="#default running-header">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <!-- Ignore in default mode -->
  </xsl:template>
  
  <xsl:template match="xhtml:aside | aside" mode="make-footnote">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <wp:fn>
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>        
      </xsl:apply-templates>
    </wp:fn>
  </xsl:template>
    
  <xsl:template match="xhtml:section/xhtml:header | section/header" mode="#default">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <wp:p>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:p>
  </xsl:template>
  
  <xsl:template match="xhtml:section/xhtml:header | section/header" mode="running-header">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <wp:p style="Header">
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:p>
  </xsl:template>
  
  <!-- ==================================
       Section and division headings
       
       ================================== -->
  
  <xsl:template match="xhtml:h1 | h1 | xhtml:h2 | h2 | xhtml:h3 | h3 | xhtml:h4 | h4 | xhtml:h5 | h5 | xhtml:h6 | h6" >
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] #default: {name(..)}/{name(.)}: "{string(.)}"</xsl:message>
    </xsl:if>
    <wp:p>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
      
    </wp:p>
  </xsl:template>
  
  <xsl:template match="
      xhtml:h1/text() | 
      xhtml:h2/text() | 
      xhtml:h3/text() | 
      xhtml:h4/text() | 
      xhtml:h5/text() | 
      xhtml:h6/text() |
      h1/text() | 
      h2/text() | 
      h3/text() | 
      h4/text() | 
      h5/text() | 
      h6/text()
      ">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <wp:run>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:call-template name="set-run-format-attributes">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <!-- Never trailing space for h1 run -->
      <xsl:value-of select="normalize-space(.)"/>
    </wp:run>
  </xsl:template>
  
  <!-- ==========================
       Figures and images
       ========================== -->
  
  <xsl:template match="xhtml:figure | figure">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:apply-templates>
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>      
  </xsl:template>
  
  <xsl:template match="xhtml:img | img">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <xsl:variable name="src" as="xs:string" select="@src"/>
    
    <xsl:variable name="base-uri" as="xs:string" select="string(base-uri(.))"/>
    
    <!-- resolve-uri() returns the input URI if it is absolute -->
    <xsl:variable name="imageUrl" as="xs:string"
      select="
      if ($src eq resolve-uri($src))
      then $src
      else relpath:newFile(relpath:getParent($base-uri), $src)"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] img: base-uri="{$base-uri}"</xsl:message>      
      <xsl:message>+ [DEBUG] img: imageUrl="{$imageUrl}"</xsl:message>
    </xsl:if>
    
    <xsl:if test="$src != ''">
      <!-- FIXME: Set style -->
      <wp:image src="{$imageUrl}">
        <!-- HTML @width and @height should work as is as long as unit is not % -->
        <xsl:sequence select="@width, @height"/>
<!--          <xsl:apply-templates select="@*"/>-->
      </wp:image>        
    </xsl:if>
    
  </xsl:template>
  
  <xsl:template match="xhtml:body/xhtml:img | body/img" priority="10">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <wp:p>
      <xsl:next-match>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:next-match>
    </wp:p>
  </xsl:template>
  
  <xsl:template match="xhtml:img/@width | xhtml:img/@height | img/@width | img/@height">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- Change unitless values to points -->
    
    <!-- Value in the XML is pixels based on 96 pixels/inch,
         so just reverse that to get inches and then points
      -->
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] image attribute: <xsl:sequence select="."/></xsl:message>
    </xsl:if>
    
    <xsl:variable name="pixels" as="xs:double" select="xs:double(.)"/>
    <xsl:variable name="inches" as="xs:double" select="$pixels div 96"/>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG]    inches: <xsl:sequence select="$inches"/></xsl:message>
    </xsl:if>
    <xsl:variable name="points" as="xs:double" select="$inches * 72"/>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG]    points: <xsl:sequence select="$points"/></xsl:message>
    </xsl:if>
    
    <xsl:attribute name="{name(.)}" select="$points"/>
    
  </xsl:template>
  
  
  <!-- ==========================
       Tables
       ========================== -->
  <xsl:template match="xhtml:table | table | xhtml:tbody | tbody | xhtml:thead | thead | xhtml:tr | tr | xhtml:col | col">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <!-- FIXME: For table rows, handle vertical spanning, which means examining preceding rows to look for
         vertically-spanning cells. Might be easiest to transform whole table into a complete matrix
         before generating the result table.
      -->
    <xsl:element name="{name(.)}" namespace="urn:ns:wordinator:simplewpml">
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>      
    </xsl:element>
    
  </xsl:template>
  
  <xsl:template match="xhtml:td | td | xhtml:th | th">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <xsl:element name="{name(.)}" namespace="urn:ns:wordinator:simplewpml">
      <xsl:call-template name="set-cell-attributes">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>      
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="xhtml:colgroup | colgroup">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <wp:cols>
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>      
    </wp:cols>
  </xsl:template>
  
  <xsl:template match="xhtml:td[empty(xhtml:p | p)] | td[empty(xhtml:p | p)] | xhtml:th[empty(xhtml:p | p)] | th[empty(xhtml:p | p)]" priority="10">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <wp:td>
      <xsl:if test="$doDebug">
        <xsl:message>+ [DEBUG] td (no paras): Calling set-cell-attributes...</xsl:message>
      </xsl:if>
      <xsl:call-template name="set-cell-attributes">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:if test="$doDebug">
        <xsl:message>+ [DEBUG] td (no paras): After set-cell-attributes.</xsl:message>
      </xsl:if>
      <wp:p>
        <xsl:apply-templates>
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>      
      </wp:p>
    </wp:td>
  </xsl:template>
  
  <xsl:template match="xhtml:td[empty(*)] | td[empty(*)] | xhtml:th[empty(*)] | th[empty(*)]" priority="15">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <wp:td>
      <xsl:call-template name="set-cell-attributes">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <wp:p>
        <xsl:call-template name="set-style">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:call-template>
        <wp:run><xsl:sequence select="normalize-space(.)"/></wp:run></wp:p>
    </wp:td>
  </xsl:template>
  
  <!-- Context should be an xhtml:td element -->
  <xsl:template name="set-cell-attributes">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-cell-attributes: Starting...</xsl:message>
    </xsl:if>
    
    <!-- NOTE: Last attribute with a given name wins -->
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-cell-attributes: Applying templates in mode set-cell-attributes to ancestor table...</xsl:message>
    </xsl:if>
    <xsl:apply-templates select="ancestor::xhtml:table[1] | ancestor::table[1]" mode="set-cell-attributes">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-cell-attributes: Calling set-cell-attributes-from-class on myself...</xsl:message>
    </xsl:if>
    <xsl:call-template name="set-cell-attributes-from-class">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:call-template>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-cell-attributes: Applying templates in mode set-cell-attributes to my attributes...</xsl:message>
    </xsl:if>
    <xsl:apply-templates select="@*" mode="set-cell-attributes">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template mode="set-cell-attributes" match="xhtml:table | table" name="set-cell-attributes-from-class">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- Default for all table cells is centered text (see Tables_Common.css) -->
    <xsl:attribute name="align" select="'center'"/>
    
    <xsl:variable name="class-values" as="xs:string*" select="tokenize(@class, ' ')"/>
    <xsl:choose>
      <xsl:when test="$class-values = ('center')">
        <xsl:attribute name="align" select="'center'"/>
      </xsl:when>
      <xsl:when test="$class-values = ('left')">
        <xsl:attribute name="align" select="'left'"/>
      </xsl:when>
      <xsl:when test="$class-values = ('right')">
        <xsl:attribute name="align" select="'right'"/>
      </xsl:when>
      <xsl:when test="$class-values = ('both', 'justify')">
        <xsl:attribute name="align" select="'both'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template mode="set-cell-attributes" match="@style[contains(., 'text-align')]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <xsl:variable name="properties" as="xs:string*" select="tokenize(., ';')"/>
    <xsl:variable name="text-align-prop" as="xs:string?" select="$properties[starts-with(., 'text-align:')]"/>
    <xsl:if test="exists($text-align-prop)">
      <xsl:variable name="align-value" as="xs:string?" select="normalize-space(tokenize($text-align-prop, ':')[2])"/>
      <xsl:attribute name="align" select="$align-value"/>
    </xsl:if>
    
  </xsl:template>
  
  <xsl:template mode="set-cell-attributes" match="@colspan | @rowspan">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:sequence select="."/>
  </xsl:template>
  
  <xsl:template mode="set-cell-attributes" match="@*" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
  </xsl:template>
  
  <!-- ==========================
       Paragraphs and spans
       ========================== -->
  
  <xsl:template match="
    xhtml:p | p | 
    xhtml:dt | dt | 
    xhtml:dd[empty(xhtml:p)] | dd[empty(p)] | 
    xhtml:pre | pre">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <wp:p>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:apply-templates select="." mode="text-before">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
      <xsl:apply-templates>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="." mode="text-after">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:p>
  </xsl:template>
  
  <xsl:template priority="10"
    match="xhtml:div/text() | div/text()"
  >
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:choose>
      <xsl:when test="normalize-space(.) = ''">
        <!-- No run -->
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match>
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:next-match>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template 
    match="
      xhtml:span/text() | 
      xhtml:dfn/text() |
      xhtml:a//text() |
      xhtml:pre//text() |
      xhtml:li//text() |
      xhtml:dt//text() |
      xhtml:dd//text() |
      xhtml:code/text() |
      xhtml:i/text() |
      xhtml:b/text() |
      xhtml:u/text() |
      xhtml:tt/text() |
      span/text() | 
      dfn/text() |
      a//text() |
      pre//text() |
      li//text() |
      dt//text() |
      dd//text() |
      code/text() |
      i/text() |
      b/text() |
      u/text() |
      tt/text()
      ">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <wp:run>
      <xsl:call-template name="set-run-format-attributes">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
           
      <xsl:call-template name="construct-run-text">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
    </wp:run>
  </xsl:template>
  
  <!-- Context should be a text() node -->
  <xsl:template name="construct-run-text">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] construct-run-text: Starting: "{.}"....</xsl:message>
    </xsl:if>
    
    <!-- Add trailing space unconditionally if there was any trailing
         space in the full text, otherwise add it conditionally (but 
         don't add it twice).
         
         If the text node is only whitespace then no text at all
         (this may or may not be appropriate; may need more sophistication
          in how we make this choice).
         
         Never adding leading space, even when in the original as that seems
         to add unwanted and unnecessary space.
      -->
    
    <xsl:variable name="base-text" as="xs:string"
      select="."
    />
    <xsl:choose>
      <xsl:when test="matches(., '^\s+$')">
        <xsl:value-of select="''"/><!-- No space -->
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="trailing-space" as="xs:string?"
          select="if (matches($base-text, '\s+$')) then ' ' else ''"
        />
        <xsl:if test="preceding-sibling::node()/self::* and matches($base-text, '^\s+')">
          <xsl:value-of select="' '"/>
        </xsl:if>
        <xsl:value-of select="normalize-space(.)"/>   
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  
  <!-- ==========================
       Lists
       ========================== -->
  
  <xsl:template match="xhtml:li[empty(xhtml:p)] | li[empty(p)]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    
    <wp:p>
      <xsl:call-template name="set-style">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>
      <xsl:apply-templates select="node() except (xhtml:ul, xhtml:ol, ul, ol)">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:p>
    
    <xsl:apply-templates select="xhtml:ul | xhtml:ol | ul | ol">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <!-- For now, assume there is no text between paragraphs in the list item, even though
       this is not a requirement of HTML markup.
       
       FIXME: Provide a more general way of handing a mix of text nodes and block nodes.
    -->
  <xsl:template match="xhtml:li[xhtml:p] | li[p]">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:apply-templates select="* except (xhtml:ul, xhtml:ol | ul | ol)">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
    <xsl:apply-templates select="xhtml:ul | xhtml:ol | ul | ol">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <!-- ==========================
       Inline things
       ========================== -->
  
  <xsl:template match="xhtml:br | br">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] xhtml:br: {name(../..)}/{name(..)}/{name(.)}</xsl:message>      
    </xsl:if>
    
    <wp:run><wp:break type="line"/></wp:run>
  </xsl:template>
  
  <xsl:template match="xhtml:br[not(ancestor::xhtml:p|ancestor::xhtml:dd)] | br[not(ancestor::p|ancestor::dd)]" priority="10">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <wp:p>
      <xsl:next-match>
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:next-match>
    </wp:p>
  </xsl:template>
    
  <!-- ==========================
       Fallback templates
       ========================== -->
  
  <xsl:template match="xhtml:p//text() | xhtml:td//text() | p//text() | td//text()" priority="-0.5">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] p//text() | td//text() fallback template: "{normalize-space(.)}"</xsl:message>
    </xsl:if>
    
    <xsl:variable name="text">
      <xsl:call-template name="construct-run-text">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:call-template>      
    </xsl:variable>
    
    <xsl:if test="exists($text) and $text != ''">
      <wp:run>
        <xsl:call-template name="set-style">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:call-template>
        <xsl:value-of select="$text"/>
      </wp:run>
    </xsl:if>  
    
    
  </xsl:template>

  <xsl:template match="text()" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- Ignore most text -->
  </xsl:template>
  
  <!-- ==========================
       Named templates
       ========================== -->
  
  <xsl:template name="set-style">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <xsl:apply-templates select="." mode="set-style">
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template mode="set-style" as="attribute()*" match="*">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="style-name" as="xs:string?">
      <xsl:apply-templates select="." mode="get-style-name">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </xsl:variable>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-style: {name(..)}/{name(.)}: style-name="{$style-name}"</xsl:message>
    </xsl:if>
    <xsl:if test="exists($style-name)">
      <xsl:attribute name="style" select="$style-name"/>
      <!-- FIXME: This will work for most cases but should really have an explicit mapping. -->
      <xsl:attribute name="styleId" select="translate($style-name, ' ', '')"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template mode="set-style" as="attribute()?" match="text()">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:variable name="style-name" as="xs:string?">
      <xsl:apply-templates select=".." mode="get-style-name">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </xsl:variable>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] set-style: {name(../..)}/{name(..)}/text(): style-name="{$style-name}"</xsl:message>
    </xsl:if>
    
    <xsl:if test="exists($style-name)">
      <xsl:attribute name="style" select="$style-name"/>
    </xsl:if>
  </xsl:template>
  
  <!-- ==============================
       Text before and text after 
       modes
       ============================== -->
  
  <xsl:template mode="text-before text-after" match="xhtml:p | p" priority="20">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] text-before|text-after: {name(..)}/{name(.)}, class="{@class}"...</xsl:message>
    </xsl:if>    
    <xsl:next-match>
      <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:template mode="text-after text-before" match="*" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- No text before or after -->
  </xsl:template>
  
  <!-- ========================================================
       Make section headers and footers 
       ======================================================== -->
       
<xsl:template mode="make-section-header-and-footer" match="*">
  <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
  
  <xsl:if test="$doDebug">
    <xsl:message>+ [DEBUG] make-section-header-and-footer: {name(..)}/{name(.)} - generating wp:headers-and-footers...</xsl:message>
  </xsl:if>
  
  <wp:headers-and-footers>
    <wp:header>
      <xsl:apply-templates select="xhtml:header" mode="running-header">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
      </xsl:apply-templates>
    </wp:header>
    <wp:footer>
      <wp:p style="Footer">
        <wp:page-number-ref format="numberInDash"/>
      </wp:p>
    </wp:footer>
  </wp:headers-and-footers>
  
</xsl:template>
  
  <!-- ========================================================
       Local functions.
       
       ======================================================== -->
  
  <!-- Get the value part of a CSS property specification.
       @param property-string A string consisting of a property name, a colon (':'), and a property value. 
                              May end with a semicolon ';'
       @return The property value as a string.
    -->
  <xsl:function name="local:get-style-value" as="xs:string">
    <xsl:param name="property-string"/>
    <xsl:variable name="base-result" as="xs:string" 
      select="normalize-space(tokenize($property-string, ':')[2])"
    />
    <xsl:variable name="result" as="xs:string"
      select="
        if (ends-with($base-result, ';')) 
        then substring($base-result, 1, string-length($base-result) - 1) 
        else $base-result"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <!-- Construct the filename for SWPX result files.
       @param context Element that is the root of the result chunk.
       @param chunk-level The chunking level in effect: "chapter" or "section"
       @param outdir The output directory the files should go to.
       @return The URI to use for the result document
    -->
  <xsl:function name="local:get-result-uri" as="xs:string">
    <xsl:param name="context" as="element()"/>
    <xsl:param name="chunk-level" as="xs:string"/>
    <xsl:param name="outdir" as="xs:string"/>
    
    <xsl:variable name="doDebug" as="xs:boolean" select="false()"/>
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] get-result-uri(): Context = {$context/xhtml:header/xhtml:h1}</xsl:message>
      <xsl:message>+ [DEBUG] get-result-uri():   count={count($context/preceding::xhtml:section[local:is-chunk(.)] |
                                                        $context/ancestor-or-self::*[local:is-chunk(.)])}</xsl:message>
    </xsl:if>   
    <xsl:variable name="filename" as="xs:string">
    <xsl:choose>
      <xsl:when test="$chunk-level = 'root'">
        <xsl:sequence select="relpath:getNamePart(base-uri($context))"/>
      </xsl:when>
      <xsl:otherwise>
        <!-- Placeholder chunk-specific filename generation -->
        <xsl:apply-templates select="$context" mode="get-chunk-filename">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
    </xsl:variable>
    
    <xsl:variable name="result" as="xs:string"
      select="relpath:newFile($outdir, $filename || '.swpx')"
    />
    <xsl:if test="$doDebug">
      <xsl:message>+ [DEBUG] get-result-uri(): returning="{$result}"</xsl:message>
    </xsl:if>
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:template mode="get-chunk-filename" match="*" as="xs:string">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>

    <xsl:variable name="chunknum" as="xs:integer"
      select="count(
      ./preceding::xhtml:section[local:is-chunk(.)] |
      ./preceding::section[local:is-chunk(.)] |
      ./ancestor-or-self::*[local:is-chunk(.)])"
    />
    <xsl:variable name="result" as="xs:string"
      select="'Chapter_' || format-number($chunknum, '00')"
    />        
    <xsl:sequence select="$result"/>
  </xsl:template>
  
  <!-- Returns true if the context element is the root of a result chunk.
       @param context Element to evaluate
       @return True if the element is a chunk for the specified chunking level setting.
    -->
  <xsl:function name="local:is-chunk" as="xs:boolean">
    <xsl:param name="context" as="element()"/>
    
    <xsl:variable name="chunk-indicator" as="xs:boolean">
      <xsl:apply-templates select="$context" mode="is-chunk">
        <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
      </xsl:apply-templates>
    </xsl:variable>
    
    <xsl:variable name="result" as="xs:boolean"
      select="$chunk-indicator"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:template mode="is-chunk" match="/*" as="xs:boolean">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    <xsl:variable name="result" as="xs:boolean" select="$chunk-level = ('root')"/>
    <xsl:sequence select="$result"/>    
  </xsl:template>
  
  <xsl:template mode="is-chunk" match="*" as="xs:boolean" priority="-1">
    <xsl:param name="doDebug" as="xs:boolean" tunnel="yes" select="false()"/>
    
    <!-- Override this mode to implement different chunking behaviors -->
    <xsl:variable name="result" as="xs:boolean" select="false()"/>
    <xsl:sequence select="$result"/>    
  </xsl:template>
  
</xsl:stylesheet>