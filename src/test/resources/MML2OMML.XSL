<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0"
        xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
        xmlns:mml="http://www.w3.org/1998/Math/MathML"
	xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        >

  <!-- this is not a real MathML to OOXML stylesheet. we always output
       the exact same thing -->
  <xsl:template match="/">
    <m:oMath><m:sSub><m:e><m:r><m:t>F</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:nor/></m:rPr><m:t>T</m:t></m:r></m:sub></m:sSub><m:r><m:t>=</m:t></m:r><m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num><m:r><m:t>1,5×</m:t></m:r><m:sSub><m:e><m:r><m:t>p</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:nor/></m:rPr><m:t>T</m:t></m:r></m:sub></m:sSub><m:r><m:t>×π×</m:t></m:r><m:sSup><m:e><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>d</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:nor/></m:rPr><m:t>n</m:t></m:r></m:sub></m:sSub><m:r><m:t>−</m:t></m:r><m:sSub><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:nor/></m:rPr><m:t>m</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:num><m:den><m:r><m:t>20</m:t></m:r></m:den></m:f></m:oMath>
  </xsl:template>
      
</xsl:stylesheet>
