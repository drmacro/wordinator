<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  version="3.0"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:relpath="http://dita2indesign/functions/relpath"
  xmlns:local="urn:localfunctions"
  exclude-result-prefixes="relpath xs local"  
  >
  <!-- =====================================================================
       DITA Community: Relative Path XSLT Functions
    
        Copyright (c) 2008, 2011 DITA2InDesign Project
        Copyright (c) 2011, 2014 DITA for Publishers
        Copyright (c) 2014 DITA Community Project
    
        This module is licensed with the Apache 2 license. It may be used
        for commercial or non-commercial purposes as long as this
        copyright is maintained.
        
        Author: W. Eliot Kimber, drmacro@yahoo.com
    
    =====================================================================-->
  
  
  <xsl:variable name="allones" select="(1,1,1,1, 1,1,1,1)" as="xs:integer*"/>
  <xsl:variable name="allzeros" select="(0,0,0,0, 0,0,0,0)" as="xs:integer*"/>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Returns the base URI of the specified context node. Fixes "file:///" to "file:/" 
      for compatibility with processors that choke on "file:///".</xd:p>
    </xd:desc>
    <xd:param name="context"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:base-uri" as="xs:string">
    <xsl:param name="context" as="node()"/>
    <xsl:variable name="baseUri" select="string(base-uri($context))" as="xs:string"/>
    <xsl:variable name="resultBaseUri" 
      select="if (starts-with($baseUri, 'file:///'))
                 then (concat('file:/', substring-after($baseUri, 'file:///')))
                 else $baseUri
      " 
      as="xs:string"/>
    <xsl:sequence select="$resultBaseUri"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Encodes a URI string.</xd:p>
    </xd:desc>
    <xd:param name="inUriString"></xd:param>
    <xd:return>Encoded URI</xd:return>
  </xd:doc>
  <xsl:function name="relpath:encodeUri" as="xs:string">
    <xsl:param name="inUriString" as="xs:string"/>
    <xsl:variable name="parts" select="tokenize($inUriString, '#')"/>
    <xsl:variable name="pathTokens" as="xs:string*">
      <xsl:for-each select="tokenize($parts[1], '/')">
        <xsl:sequence select="encode-for-uri(.)"/>  
      </xsl:for-each>      
    </xsl:variable>
    <xsl:variable name="escapedFragId" as="xs:string"
      select="if ($parts[2]) then concat('#', encode-for-uri($parts[2])) else ''"
    />
    <xsl:variable name="result" as="xs:string"
      select="concat(string-join($pathTokens, '/'), $escapedFragId)"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Unecodes an encoded URI string encoded in UTF-8 encoding.</xd:p>
    </xd:desc>
    <xd:param name="inUriString"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:unencodeUri" as="xs:string">
    <xsl:param name="inUriString" as="xs:string"/>
    <xsl:variable name="firstChar" select="substring($inUriString, 1, 1)" as="xs:string"/>
    <xsl:variable name="outUriString" as="xs:string*">
      <xsl:for-each select="tokenize($inUriString, '/')">
        <xsl:variable name="pathComponent" select="." as="xs:string"/>
        <xsl:variable name="unencodedComponent" select="relpath:unencodeString(.)" as="xs:string"/>
        <xsl:sequence select="$unencodedComponent"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="pathComponentsJoined" select="string-join($outUriString, '/')"/>
    <xsl:sequence
      select="$pathComponentsJoined"
    />
  </xsl:function>
  
  <xsl:function name="relpath:unencodeString" as="xs:string">
    <xsl:param name="inString" as="xs:string"/>
    <xsl:variable name="tokens" as="xs:string*">
      <xsl:analyze-string select="$inString" regex="%[0-9A-Fa-f][0-9A-Fa-f]">
        <xsl:matching-substring>
          <xsl:sequence select="."/>
        </xsl:matching-substring>
        <xsl:non-matching-substring>
          <xsl:sequence select="."/>
        </xsl:non-matching-substring>
      </xsl:analyze-string>
    </xsl:variable>
<!--    <xsl:message>+ [DEBUG] unencodeString(): tokens=<xsl:sequence select="string-join($tokens, '|')"/></xsl:message>-->
    <xsl:variable name="resultTokens" as="xs:string*"
      select="relpath:unencodeStringTokens($tokens, ())"
    />
    <xsl:sequence select="string-join($resultTokens, '')"/>
  </xsl:function>
  
  <xsl:function name="relpath:unencodeStringTokens" as="xs:string*">
    <xsl:param name="tokens" as="xs:string*"/>
    <xsl:param name="resultTokens" as="xs:string*"/> 
    <xsl:choose>
      <xsl:when test="count($tokens) = 0">
        <xsl:sequence select="$resultTokens"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="token" select="$tokens[1]" as="xs:string"/>
        <xsl:choose>
          <xsl:when test="starts-with($token, '%')">
            <xsl:sequence select="relpath:processEncodedChars($token, $tokens[position() > 1], $resultTokens)"/>
          </xsl:when>
          <xsl:otherwise>
            <!-- Reset character code and copy token to result token list: -->
            <xsl:sequence select="relpath:unencodeStringTokens($tokens[position() > 1], ($resultTokens, $tokens[1]))"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>   
  </xsl:function>
  
  <xsl:function name="relpath:processEncodedChars" as="xs:string*">
    <xsl:param name="byte1Str" as="xs:string"/>
    <xsl:param name="tokens" as="xs:string*"/>
    <xsl:param name="resultTokens" as="xs:string*"/>
     
    <xsl:variable name="byte1" as="xs:integer"
      select="relpath:hex-to-char($byte1Str)"
    /> 
    <xsl:variable name="sequenceLength" as="xs:integer"
    >
      <!-- The number of leading bits for the first byte
           indicate the number of bytes for the encoded
           character.
        -->
      <xsl:choose>
        <xsl:when test="$byte1 ge 248">
          <xsl:sequence select="5"/><!-- 11111xxx -->
        </xsl:when>
        <xsl:when test="$byte1 ge 240">
          <xsl:sequence select="4"/><!-- 1111xxxx -->
        </xsl:when>
        <xsl:when test="$byte1 ge 224">
          <xsl:sequence select="3"/><!-- 111xxxxx -->
        </xsl:when>
        <xsl:when test="$byte1 ge 192">
          <xsl:sequence select="2"/><!-- 11xxxxxx -->
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="1"/>
        </xsl:otherwise>
      </xsl:choose>      
    </xsl:variable>
    
    <!-- Take byte1 and mask out leading bits that indicate sequence length -->
    <xsl:variable name="byte1bits" as="xs:integer">
      <xsl:choose>
        <xsl:when test="$sequenceLength = 1">
          <xsl:sequence select="$byte1"/>
        </xsl:when>
        <xsl:when test="$sequenceLength = 2">
          <xsl:sequence select="$byte1 - 192"/>
        </xsl:when>
        <xsl:when test="$sequenceLength = 3">
          <xsl:sequence select="$byte1 - 224"/>
        </xsl:when>
        <xsl:when test="$sequenceLength = 4">
          <xsl:sequence select="$byte1 - 240"/>
        </xsl:when>
        <xsl:when test="$sequenceLength = 5">
          <xsl:sequence select="$byte1 - 248"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="yes">+ [ERROR] Sequence length <xsl:sequence select="$sequenceLength"/> is not a good value. Must be between 1 and 5.</xsl:message>
        </xsl:otherwise>
      </xsl:choose>      
    </xsl:variable>
    <xsl:variable name="bitmask" 
      select="$allones[position() le $sequenceLength],
              $allzeros[position() gt $sequenceLength]"></xsl:variable>
    <!-- Shift byte1 bits to left -->
    <xsl:variable name="byte1bits" 
      select="relpath:getBits($byte1)" 
      as="xs:integer*"/>
    <xsl:variable name="accumulatedBits" select="relpath:bitwiseXor($byte1bits, $bitmask)"/>
    <xsl:variable name="characterBits" as="xs:integer*"
      select="relpath:decodeRemainingUtfBytes($tokens[position() lt $sequenceLength], $accumulatedBits)"
    />
    <xsl:variable name="char" select="codepoints-to-string(relpath:bits2int($characterBits))" as="xs:string"/>
    <xsl:sequence select="relpath:unencodeStringTokens($tokens[position() ge $sequenceLength], ($resultTokens, $char))"/>
  </xsl:function>

  <xsl:function name="relpath:decodeRemainingUtfBytes" as="xs:integer*">
    <xsl:param name="tokens" as="xs:string*"/>
    <xsl:param name="accumulatedBits" as="xs:integer*"/>
    <xsl:variable name="result" as="xs:integer*">
      <xsl:choose>
        <xsl:when test="count($tokens) = 0">
          <xsl:sequence select="$accumulatedBits"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="byte" as="xs:integer" 
            select="relpath:hex-to-char($tokens[1])"/>
          <!-- Subsequence bytes are 10xxxxxx, so mask out leading bit -->
          <xsl:variable name="bitList" as="xs:integer*"
            select="relpath:getBits($byte - 128)[position() gt 2]"
          />
          <xsl:sequence 
            select="relpath:decodeRemainingUtfBytes($tokens[position() gt 1], ($accumulatedBits, $bitList))"/>
        </xsl:otherwise>
      </xsl:choose>      
    </xsl:variable>
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:getBits" as="xs:integer*">
    <xsl:param name="byte" as="xs:integer"/>
    <xsl:variable name="bitSeq" as="xs:integer*">
      <xsl:sequence select="local:doGetBits($byte, 7, ())"/>
    </xsl:variable>
    <xsl:sequence select="$bitSeq"/>
  </xsl:function>
  
  <xsl:function name="local:doGetBits" as="xs:integer*">
    <xsl:param name="startingValue" as="xs:integer"/>
    <xsl:param name="bitpos" as="xs:integer"/>
    <xsl:param name="bits" as="xs:integer*"/>
    <xsl:variable name="powerOfTwo" as="xs:integer" 
      select="relpath:exp(2, $bitpos)"/>
    <xsl:variable name="testValue" as="xs:integer"
      select="if ($powerOfTwo gt 0) 
         then ($startingValue idiv $powerOfTwo) 
         else $startingValue"
    />
    <xsl:variable name="bit" as="xs:integer"
      select="if ($testValue gt 0) then 1 else 0"
    />
    <xsl:choose>
      <xsl:when test="$bitpos = 0">
        <xsl:sequence select="$bits, $bit"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="remainingValue" as="xs:integer"
          select="if ($startingValue ge $powerOfTwo) 
                     then $startingValue - $powerOfTwo 
                     else $startingValue"
        />
        <xsl:sequence 
          select="local:doGetBits($remainingValue, $bitpos - 1, ($bits, $bit))"/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:function>
  
  <xsl:function name="relpath:exp" as="xs:integer">
    <xsl:param name="base" as="xs:integer"/>
    <xsl:param name="exponent" as="xs:integer"/>
    <xsl:variable name="result" as="xs:integer"
      select="local:doExponentiation($base, $exponent, $base)"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="local:doExponentiation" as="xs:integer">
    <xsl:param name="base" as="xs:integer"/>
    <xsl:param name="exponent" as="xs:integer"/>
    <xsl:param name="accumulatedValue" as="xs:integer"/>
    <xsl:choose>
      <xsl:when test="$exponent = 0">
        <xsl:sequence select="0"/>
      </xsl:when>
      <xsl:when test="$exponent = 1">
        <xsl:sequence select="$accumulatedValue"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="local:doExponentiation($base, $exponent - 1, $accumulatedValue * $base)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="relpath:shiftLeft" as="xs:integer">
    <xsl:param name="bits" as="xs:integer*"/>
    <xsl:param name="shiftSize" as="xs:integer"/>
    <xsl:variable name="bitValue" as="xs:integer"
      select="relpath:bits2int($bits)"
    />
    <xsl:variable name="result" as="xs:integer"
      select="$bitValue * relpath:exp(2, $shiftSize)"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:bits2int" as="xs:integer">
    <xsl:param name="bits" as="xs:integer*"/>
    <xsl:variable name="result" as="xs:integer"
         select="local:doBits2int($bits, 0)"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="local:doBits2int" as="xs:integer">
    <xsl:param name="bits" as="xs:integer*"/>
    <xsl:param name="accumulatedValue" as="xs:integer"/>
    <xsl:choose>
      <xsl:when test="count($bits) = 0">
        <xsl:sequence select="$accumulatedValue"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="bitValue" as="xs:integer">
          <xsl:choose>
            <xsl:when test="count($bits) gt 1">
              <xsl:sequence select="$bits[1] * relpath:exp(2, count($bits) - 1)"/>    
            </xsl:when>
            <xsl:otherwise>
              <xsl:sequence select="$bits[1]"/>
            </xsl:otherwise>
          </xsl:choose>          
        </xsl:variable>
        
        <xsl:sequence select="local:doBits2int($bits[position() gt 1], $accumulatedValue + $bitValue)"/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:function>
  
  <xsl:function name="relpath:bitwiseAnd" as="xs:integer*">
    <xsl:param name="byte1" as="xs:integer*"/>
    <xsl:param name="byte2" as="xs:integer*"/>
    <xsl:if test="count($byte1) ne count($byte2)">
      <xsl:message terminate="yes"> - [ERROR] relpath:bitwiseAnd(): Bit count not equal. Byte 1 is <xsl:sequence select="count($byte1)"/>, byte 2 is <xsl:sequence select="count($byte2)"/>.</xsl:message>
    </xsl:if>
   
   <xsl:variable name="result" as="xs:integer*">
     <xsl:for-each select="$byte1">
       <xsl:variable name="bitPos" as="xs:integer" select="position()"/>
       <!--<xsl:message>+ [DEBUG] bitwiseAnd(): testValue=<xsl:sequence select="$testValue"/></xsl:message>-->
       <xsl:sequence select="if ((. + $byte2[$bitPos]) gt 1) then 1 else 0"/>
     </xsl:for-each>
   </xsl:variable>
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:bitwiseXor" as="xs:integer*">
    <xsl:param name="byte1" as="xs:integer*"/>
    <xsl:param name="byte2" as="xs:integer*"/>
    <xsl:if test="count($byte1) ne count($byte2)">
      <xsl:message terminate="yes"> - [ERROR] relpath:bitwiseXor(): Bit count not equal. Byte 1 is <xsl:sequence select="count($byte1)"/>, byte 2 is <xsl:sequence select="count($byte2)"/>.</xsl:message>
    </xsl:if>
    
    <xsl:variable name="result" as="xs:integer*">
      <xsl:for-each select="$byte1">
        <xsl:variable name="bitPos" as="xs:integer" select="position()"/>
        <!--<xsl:message>+ [DEBUG] bitwiseAnd(): testValue=<xsl:sequence select="$testValue"/></xsl:message>-->
        <xsl:sequence select="if ((. + $byte2[$bitPos]) gt 1) then 0 else ."/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:unescapeUriCharacter" as="xs:string">
    <xsl:param name="escapedCharStr" as="xs:string"/>
    <xsl:variable name="cadr" select="substring($escapedCharStr, 2)"/>
    <xsl:variable name="resultChar" as="xs:string">
      <!-- cadr must be pair of hex digits -->
      <xsl:variable name="codepoint" select="relpath:hex-to-char($cadr)" as="xs:integer"/>
      <xsl:sequence select="codepoints-to-string($codepoint)"/>
    </xsl:variable>
    <xsl:sequence select="$resultChar"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Gets the filename of the resource (last path component) as for Java File.getName().</xd:p>
    </xd:desc>
    <xd:param name="sourcePath"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getName" as="xs:string">
    <!-- As for Java File.getName(): returns the last
         component of the path.
    -->
    <xsl:param name="sourcePath" as="xs:string"/>
    <xsl:value-of select="tokenize($sourcePath, '/')[last()]"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Gets the name part of the resource filename, that is, the filename with any extension removed.</xd:p>
    </xd:desc>
    <xd:param name="sourcePath"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getNamePart" as="xs:string">
    <!-- Returns the name part of a filename, excluding
         any trailing extension.
    -->
    <xsl:param name="sourcePath" as="xs:string"/>
    <xsl:variable name="fullName" select="relpath:getName($sourcePath)" as="xs:string"/>
    <xsl:variable name="result" as="xs:string"
      select="if (contains($fullName, '.'))
                 then string-join(tokenize($fullName, '\.')[position() &lt; last()], '.')
                 else $fullName
             "
    />
    <xsl:value-of select="$result"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Gets the extension, if any, of the resource's filename.</xd:p>
    </xd:desc>
    <xd:param name="sourcePath"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getExtension" as="xs:string">
    <!-- Returns the extension part of a filename, excluding the
         leading "."
    -->
    <xsl:param name="sourcePath" as="xs:string"/>
    <xsl:variable name="fullName" select="relpath:getName($sourcePath)" as="xs:string"/>
    <xsl:variable name="result" as="xs:string"
      select="if (contains($fullName, '.'))
      then tokenize($fullName, '\.')[last()]
      else ''
      "
    />
    <xsl:value-of select="$result"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>As for Java File.getParent(): returns all but the last
        components of the path.</xd:p>
    </xd:desc>
    <xd:param name="sourcePath"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getParent" as="xs:string">
    <xsl:param name="sourcePath" as="xs:string"/>
    <xsl:value-of select="string-join(tokenize($sourcePath, '/')[position() &lt; last()], '/')"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>As for Java File(File, path)): Returns a new a absolute path representing
        the new file. File must be a path (because XSLT has no way to distinguish
        a file from a directory).</xd:p>
    </xd:desc>
    <xd:param name="parentPath">Parent directory for the new file.</xd:param>
    <xd:param name="childFile">Path and name of child file. May be an absolute or relative
    path.</xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:newFile" as="xs:string">
    <xsl:param name="parentPath" as="xs:string"/>
    <xsl:param name="childFile" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="matches($childFile, '^(/|[a-z]+:)')">
        <xsl:value-of select="$childFile"/>
      </xsl:when>
      <xsl:when test="$parentPath = '/'">
        <xsl:variable name="tempPath" select="concat($parentPath, $childFile)" as="xs:string"/>
        <xsl:variable name="result"
          select="relpath:getAbsolutePath($tempPath)"/>
        <xsl:value-of select="$result"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="tempParentPath" as="xs:string" 
          select="if (ends-with($parentPath, '/') and $parentPath != '/') 
          then substring($parentPath, 1, string-length($parentPath) - 1) 
          else $parentPath" />
        <xsl:variable name="parentTokens" select="tokenize($tempParentPath, '/')" as="xs:string*"/>
        <xsl:variable name="firstToken" select="$parentTokens[1]" as="xs:string?"/>
        <xsl:variable name="childTokens" select="tokenize($childFile, '/')" as="xs:string*"/>
        <xsl:variable name="tempPath" select="string-join(($parentTokens, $childTokens), '/')" as="xs:string"/>
        <xsl:variable name="result"
          select="if ($firstToken = '..') 
              then $tempPath 
              else relpath:getAbsolutePath($tempPath)"/>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="relpath:getResourcePartOfUri" as="xs:string">
    <xsl:param name="uriString" as="xs:string"/>
    <xsl:variable name="resourcePart" as="xs:string"
      select="
      if (contains($uriString, '#')) 
      then substring-before($uriString, '#') 
      else normalize-space($uriString)
      "/>
    <xsl:sequence select="$resourcePart"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Given a path resolves any ".." or "." terms to produce an absolute path</xd:p>
    </xd:desc>
    <xd:param name="sourcePath"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getAbsolutePath" as="xs:string">
    <xsl:param name="sourcePath" as="xs:string"/>
    <xsl:variable name="pathTokens" select="tokenize($sourcePath, '/')" as="xs:string*"/>
    <xsl:if test="false()">
      <xsl:message> + DEBUG relpath:getAbsolutePath(): Starting</xsl:message>
      <xsl:message> +       sourcePath="<xsl:value-of select="$sourcePath"/>"</xsl:message>
    </xsl:if>
    <xsl:variable name="baseResult" 
    select="string-join(relpath:makePathAbsolute($pathTokens, ()), '/')" as="xs:string"/>
    <xsl:variable name="baseResult2" 
      select="if (ends-with($baseResult, '/')) 
                 then substring($baseResult, 1, string-length($baseResult) -1) 
                 else $baseResult" as="xs:string"/>
    <xsl:variable name="result" as="xs:string"
       select="if (starts-with($sourcePath, '/') and not(starts-with($baseResult2, '/')))
                  then concat('/', $baseResult2)
                  else $baseResult2
               "
    />
    <xsl:if test="false()">
      <xsl:message> + DEBUG: result="<xsl:value-of select="$result"/>"</xsl:message>
    </xsl:if>
    <xsl:value-of select="$result"/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Removes any relative components from the path.</xd:p>
    </xd:desc>
    <xd:param name="pathTokens"></xd:param>
    <xd:param name="resultTokens"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:makePathAbsolute" as="xs:string*">
    <xsl:param name="pathTokens" as="xs:string*"/>
    <xsl:param name="resultTokens" as="xs:string*"/>
    <xsl:if test="false()">
      <xsl:message> + DEBUG: relpath:makePathAbsolute(): Starting...</xsl:message>
      <xsl:message> + DEBUG:    pathTokens="<xsl:value-of select="string-join($pathTokens, ',')"/>"</xsl:message>
      <xsl:message> + DEBUG:    resultTokens="<xsl:value-of select="string-join($resultTokens, ',')"/>"</xsl:message>
    </xsl:if>
    <xsl:sequence select="if (count($pathTokens) = 0)
                             then $resultTokens
                             else if ($pathTokens[1] = '.')
                                  then relpath:makePathAbsolute($pathTokens[position() > 1], $resultTokens)
                                  else if ($pathTokens[1] = '..')
                                       then relpath:makePathAbsolute($pathTokens[position() > 1], $resultTokens[position() &lt; last()])
                                       else relpath:makePathAbsolute($pathTokens[position() > 1], ($resultTokens, $pathTokens[1]))
                         "/>
  </xsl:function>
  
  <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Calculate relative path that gets from from source path to target path.</xd:p>
      <xd:pre>
        Given:
        
        [1]  Target: /A/B/C
        Source: /A/B/C/X
        
        Return: "X"
        
        [2]  Target: /A/B/C
        Source: /E/F/G/X
        
        Return: "/E/F/G/X"
        
        [3]  Target: /A/B/C
        Source: /A/D/E/X
        
        Return: "../../D/E/X"
        
        [4]  Target: /A/B/C
        Source: /A/X
        
        Return: "../../X"
      </xd:pre>
    </xd:desc>
    <xd:param name="source"></xd:param>
    <xd:param name="target"></xd:param>
    <xd:return></xd:return>
  </xd:doc>
  <xsl:function name="relpath:getRelativePath" as="xs:string">
    <xsl:param name="source" as="xs:string"/><!-- Path to get relative path *from* -->
    <xsl:param name="target" as="xs:string"/><!-- Path to get relataive path *to* -->
    <xsl:variable name="effectiveSource" as="xs:string"
      select="if (ends-with($source, '/') and string-length($source) > 1) then substring($source, 1, string-length($source) - 1) else $source"
    />
    <xsl:if test="false()">
      <xsl:message>+ [DEBUG]: relpath:getRelativePath(): Starting...</xsl:message>
      <xsl:message>+ [DEBUG]:     source="<xsl:value-of select="$effectiveSource"/>"</xsl:message>
      <xsl:message>+ [DEBUG]:     target="<xsl:value-of select="$target"/>"</xsl:message>
    </xsl:if>
    <xsl:variable name="sourceTokens" select="tokenize((if (starts-with($effectiveSource, '/')) then substring-after($effectiveSource, '/') else $effectiveSource), '/')" as="xs:string*"/>
    <xsl:variable name="targetTokens" select="tokenize((if (starts-with($target, '/')) then substring-after($target, '/') else $target), '/')" as="xs:string*"/>
    <xsl:choose>
      <xsl:when test="(count($sourceTokens) > 0 and count($targetTokens) > 0) and 
                      (($sourceTokens[1] != $targetTokens[1]) and 
                       (contains($sourceTokens[1], ':') or contains($targetTokens[1], ':')))">
        <!-- Must be absolute URLs with different schemes, cannot be relative, return
        target as is. -->
        <xsl:value-of select="$target"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="resultTokens" 
          select="relpath:analyzePathTokens($sourceTokens, $targetTokens, ())" as="xs:string*"/>      
        <xsl:if test="false()">
          <xsl:message>+ [DEBUG] relpath:getRelativePath(): resultTokens="<xsl:sequence select="$resultTokens"/>"</xsl:message>
        </xsl:if>
        <xsl:variable name="result" select="string-join($resultTokens, '/')" as="xs:string"/>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <!-- Convert an operating-system-specific path to
       a URL.
    -->
  <xsl:function name="relpath:toUrl" as="xs:string">
    <xsl:param name="filepath" as="xs:string"/>
    <xsl:variable name="url" as="xs:string"
      select="if (contains($filepath, '\'))
       then translate($filepath, '\', '/')
       else $filepath
      "
    />
    <!-- If the URL use a Windows path or a Unix path,
         append "file:/" otherwise use it as is.
      -->
    <xsl:variable name="fileUrl" as="xs:string"
      select="
      if (matches($url, '^[a-zA-Z]:'))
        then concat('file:/', $url)
        else if (starts-with($url, '/')) 
                then concat('file:', $url) 
                else $url
        "
    />
    <!-- Escape any spaces in the URL. This is 
         not complete but should be sufficient to
         handle all Unix cases and most Windows cases.
      -->
    <xsl:variable name="escapedUrl" 
      select="replace($fileUrl, ' ', '%20')"
    />
    <xsl:sequence select="$escapedUrl"/>
  </xsl:function>
  
  <xsl:function name="relpath:toFile" as="xs:string">
    <xsl:param name="url" as="xs:string"/>
    <xsl:param name="platform" as="xs:string"/><!-- One of "windows", "nx" -->
    <xsl:variable name="unescapedUrl" select="relpath:unencodeUri($url)" as="xs:string"/>
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] -------------</xsl:message>
      <xsl:message>+ [DEBUG] toFile(): url         ='<xsl:sequence select="$url"/>', platform='<xsl:sequence select="$platform"/>'</xsl:message>
      <xsl:message>+ [DEBUG] toFile(): unescapedUrl='<xsl:sequence select="$unescapedUrl"/>', platform='<xsl:sequence select="$platform"/>'</xsl:message>
    </xsl:if>
    <xsl:variable name="result" as="xs:string"
       select="
       if ($platform = 'windows')
          then relpath:urlToWindowsFile($unescapedUrl)
          else relpath:urlToNxFile($unescapedUrl)
       "
    />
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] toFile(): result='<xsl:sequence select="$result"/>'</xsl:message>
    </xsl:if>
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:urlToNxFile" as="xs:string">
    <xsl:param name="url" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="matches($url, '^file:.+')">
        <xsl:variable name="filepath" select="substring-after($url, 'file:')" as="xs:string"/>
        <xsl:variable name="result" as="xs:string"
          select="
          if (starts-with($filepath, '///'))
          then substring($filepath, 3)
          else $filepath
          "
        />
        <xsl:sequence select="$result"/>
      </xsl:when>
      <xsl:when test="matches($url, '^[a-zA-Z]+:.+')">
        <xsl:sequence select="'{cannot convert absolute URL to file path}'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$url"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="relpath:urlToWindowsFile" as="xs:string">
    <xsl:param name="url" as="xs:string"/>
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] urlToWindowsFile(): url='<xsl:sequence select="$url"/>'</xsl:message>
    </xsl:if>
    <xsl:variable name="basePath" as="xs:string"
       select="relpath:urlToNxFile($url)"
    />
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] urlToWindowsFile(): basePath='<xsl:sequence select="$basePath"/>'</xsl:message>
    </xsl:if>
    <xsl:variable name="windowsPath" as="xs:string" 
      select="
     replace(
      replace($basePath, '/', '\\'), 
      '%20', 
      ' ')" 
      />
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] urlToWindowsFile(): windowsPath='<xsl:sequence select="$windowsPath"/>'</xsl:message>
    </xsl:if>
    <xsl:variable name="result" as="xs:string">
      <xsl:choose>
        <xsl:when test="matches($windowsPath, '^\\[a-zA-Z]+:')">
          <xsl:sequence select="substring($windowsPath, 2)"/>
        </xsl:when>
        <xsl:when test="matches($windowsPath, '^\\\\\\[a-zA-Z]+:')">
          <xsl:sequence select="substring($windowsPath, 4)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="$windowsPath"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <xsl:if test="false()">      
      <xsl:message>+ [DEBUG] urlToWindowsFile(): Returning '<xsl:sequence select="$result"/>'</xsl:message>
    </xsl:if>
    <xsl:sequence select="$result"/>
  </xsl:function> 
  
  
  
  <xsl:function name="relpath:analyzePathTokens" as="xs:string*">
    <xsl:param name="sourceTokens" as="xs:string*"/>
    <xsl:param name="targetTokens" as="xs:string*"/>
    <xsl:param name="resultTokens" as="xs:string*"/>
    <xsl:if test="false()">
    <xsl:message> + DEBUG: relpath:analyzePathTokens(): Starting...</xsl:message>
    <xsl:message> + DEBUG:     sourceTokens=<xsl:value-of select="string-join($sourceTokens, ',')"/></xsl:message>
    <xsl:message> + DEBUG:     targetTokens=<xsl:value-of select="string-join($targetTokens, ',')"/></xsl:message>
    <xsl:message> + DEBUG:     resultTokens=<xsl:value-of select="string-join($resultTokens, ',')"/></xsl:message>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="count($sourceTokens) = 0">
        <!-- Append remaining target tokens (if any) to the result -->
        <xsl:sequence select="$resultTokens, $targetTokens"/>
      </xsl:when>
      <xsl:otherwise>
        <!-- Still source tokens, so see if source[1] = target[1] -->
        <xsl:choose>
          <!-- If they are equal, go to the next level in the paths: -->
          <xsl:when test="(count($targetTokens) > 0) and ($sourceTokens[1] = $targetTokens[1])">
            <xsl:sequence select="relpath:analyzePathTokens($sourceTokens[position() > 1], $targetTokens[position() > 1], $resultTokens)"/>
          </xsl:when>
          <xsl:otherwise>
            <!-- Paths must diverge at this point. Append one ".." for each token
            left in the source: -->
            <xsl:if test="false()">
              <xsl:message>+ [DEBUG] constructing goUps: $sourceTokens=<xsl:sequence select="$sourceTokens"/></xsl:message>
            </xsl:if>
            <xsl:variable name="goUps" as="xs:string*">
              <xsl:for-each select="$sourceTokens">
                <xsl:sequence select="'..'"/>
              </xsl:for-each>
            </xsl:variable>
            <xsl:sequence select="string-join(($resultTokens, $goUps, $targetTokens), '/')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>    
  </xsl:function>
  
  <xsl:function name="relpath:getFragmentId" as="xs:string">
    <xsl:param name="uri" as="xs:string"/>
    <!-- Returns either the empty string, if there is no fragment
         identifier, or the string following the first "#" in the 
         URI up to, but not including, any query component.
      -->
    <xsl:variable name="baseFragid" as="xs:string"
      select="if (contains($uri, '#'))
      then substring-after($uri, '#')
      else ''"
    />
    <xsl:variable name="result" as="xs:string"
      select="if (contains($baseFragid, '?')) 
      then substring-before($baseFragid, '?') 
      else $baseFragid"
    />
    <xsl:sequence select="$result"/>
  </xsl:function>
  
  <xsl:function name="relpath:hex-to-char" as="xs:integer">
    <xsl:param name="in" as="xs:string"/> <!-- e.g. 030C -->
    <xsl:sequence select="
      if (string-length($in) eq 1)
      then relpath:hex-digit-to-integer($in)
      else 16 * relpath:hex-to-char(substring($in, 1, string-length($in)-1)) +
      relpath:hex-digit-to-integer(substring($in, string-length($in)))"/>
  </xsl:function>
  
  <xsl:function name="relpath:hex-digit-to-integer" as="xs:integer">
    <xsl:param name="char" as="xs:string"/>
    <xsl:sequence 
      select="string-length(substring-before('0123456789ABCDEF',
      $char))"/>
  </xsl:function>
  
   <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl">
    <xd:desc>
      <xd:p>Prefix non-aboslute path with $relativePath</xd:p>
      <xd:pre>
        Given:
        
        [1]  
        relativePath: ../../
        filePath: A/B/C/X
        
        Return: "../../A/B/C/X"
        
         
        [2]  
        relativePath: ../../
        filePath: /A/B/C/X
        
        Return: "/A/B/C/X"
       
      </xd:pre>
    </xd:desc>
    <xd:param name="relativePath">
    	<xd:p>A relative path prefix. Example: ../</xd:p>
    </xd:param>
    <xd:param name="filePath">
    	<xd:p>a file path</xd:p>
    </xd:param>
    <xd:return></xd:return>
  </xd:doc>
    <xsl:function name="relpath:fixRelativePath" as="xs:string*">
    	<xsl:param name="relativePath" as="xs:string*"/>
    	<xsl:param name="filePath" as="xs:string*"/>
  		
     	<xsl:variable name="prefix">
  	    	<xsl:choose>
  	    		<xsl:when test="substring($filePath, 1, 1) = '/'">
  	    			<xsl:value-of select="''" />
  	    		</xsl:when>
  	    		<xsl:otherwise>
  	    			<xsl:value-of select="$relativePath" />
  	    		</xsl:otherwise>
  	    	    	
  	    	</xsl:choose>   		
    	</xsl:variable>
    	
    	<xsl:value-of select="concat($prefix, $filePath)" />
    	
    	
  </xsl:function>
  
</xsl:stylesheet>
