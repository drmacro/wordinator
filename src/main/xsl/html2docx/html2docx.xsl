<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:wp="urn:ns:wordinator:simplewpml"
  xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:local="urn:ns:local-functions"
  xmlns:relpath="http://dita2indesign/functions/relpath"
  exclude-result-prefixes="xs local relpath"
  version="3.0"
  expand-text="yes"
  >
  <!-- ========================================================================
       HTML to DOCX Transform
       
       This transform generates Simple Word Processing XML
       documents from HTML5 documents for subsequent
       processing by the Wordinator Java code to generate
       working DOCX files.
       
       NOTE: The extension used for the generated files is ".swpx" to make
       it easy to find these files in a directory for subsequent processing.
       
       Direct output is a log of the generation process.
       
       @param outdir Output directory. Default is "out/swpx" relative to
       the input document's directory.
       ======================================================================== -->
  
  <xsl:import href="../lib/relpath_util.xsl"/>
  <xsl:import href="baseProcessing.xsl"/>
  <xsl:import href="get-style-name.xsl"/>
  <xsl:import href="set-format-attributes.xsl"/>
  <xsl:import href="cleanup-swpx.xsl"/>
  
  <xsl:output method="xml"
    indent="yes"
  />
  
  <xsl:output method="xml" name="swpx"
    indent="no"    
  />
  
  <xsl:param name="outdir" as="xs:string" select="'out/swpx'"/>
  <xsl:variable name="outputDirectory" as="xs:string"
    select="relpath:newFile(relpath:getParent(string(document-uri(/))), $outdir)"
  />
  
  <!-- Controls level at which new result SWPX documents are generated, by
       default the root document. Value is a string that is used to determine
       if an element is or is not the start of a new chunk. 
       
       Enables creating multiple DOCX files from a single input document.
       
       The local:is-chunk(.) function is used to determine if a given
       element is or is not the start of a new result chunk.
    -->
  <xsl:param name="chunk-level" as="xs:string" select="'root'"/>
  
  <xsl:template match="/">
    <xsl:param name="doDebug" as="xs:boolean" select="false()"/>
    
    <log started="{format-dateTime(current-dateTime(), '[Y000]-[m0]}-[d0]')}">
      <message>+ [INFO] inputdoc={document-uri(.)}</message>
      <message>+ [INFO] outdir={$outdir}</message>
      <generation-process>
        <xsl:if test="$doDebug">
          <xsl:message>+ [DEBUG] Applying templates to root element: {namespace-uri(/*)}:{name(/*)}</xsl:message>
        </xsl:if>
        <xsl:apply-templates select="/*">
          <xsl:with-param name="doDebug" as="xs:boolean" tunnel="yes" select="$doDebug"/>
        </xsl:apply-templates>
      </generation-process>     
    </log>
  </xsl:template>
</xsl:stylesheet>