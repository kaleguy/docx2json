<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                exclude-result-prefixes="w r rels a wp cp dc dcterms mc">

    <xsl:output method="xml"
                indent="no"
                encoding="UTF-8"/>
    <xsl:strip-space elements="*"/>

    <xsl:template match="/">
        <xml>
            <head>
                <xsl:apply-templates select="w:document/cp:coreProperties"/>
            </head>
            <body>
                <xsl:apply-templates />
            </body>
        </xml>
    </xsl:template>

    <!-- document data e.g. title and description -->
    <xsl:template match="dc:* | dcterms:*">
        <xsl:element name="{local-name()}">
            <xsl:value-of select="."/>
        </xsl:element>
    </xsl:template>

    <!-- 'alternate content' is skipped -->
    <xsl:template match="mc:AlternateContent" mode="ol">

    </xsl:template>>

    <!-- TOC -->
    <xsl:template match="w:p[w:pPr/w:pStyle/@w:val[starts-with(., 'ContentsHeading')]]" mode="toc">
        <toc>
            <heading><xsl:value-of select="."/></heading>
            <links>
               <xsl:apply-templates select="//w:p" mode="toc_section"/>
            </links>
        </toc>
    </xsl:template>
    <xsl:template match="w:p[w:pPr/w:pStyle/@w:val[starts-with(., 'Contents1')]]" mode="toc_section">
       <xsl:variable name="link_text">
           <xsl:call-template name="link_text">
               <xsl:with-param name="link" select="w:hyperlink"/>
           </xsl:call-template>
       </xsl:variable>
       <link name="{$link_text}"
             target="{w:hyperlink/@w:anchor}"
             style="{w:pPr/w:pStyle/@w:val}">
       </link>
    </xsl:template>
    <xsl:template match="w:t" mode="toc_section"/>

    <!-- Alternate TOC spec -->
    <xsl:template match="w:sdt" priority="1"/>
    <xsl:template match="w:sdt[w:sdtContent/w:p/w:hyperlink]" priority="2">
        <toc>
            <xsl:apply-templates/>
        </toc>
    </xsl:template>
    <xsl:template match="w:sdtContent">
        <heading>
            <xsl:value-of select="w:p[w:pPr/w:pStyle[@w:val='ContentsHeading']]/w:r/w:t"/>
            <xsl:value-of select="w:p[w:pPr/w:pStyle[@w:val='TOCHeading']]/w:r/w:t"/>
        </heading>
        <links>
            <xsl:for-each select="w:p[w:hyperlink]">
                <xsl:variable name="link_text">
                    <xsl:call-template name="link_text">
                        <xsl:with-param name="link" select="w:hyperlink"/>
                    </xsl:call-template>
                </xsl:variable>
                <link name="{$link_text}"
                      target="{w:hyperlink/@w:anchor}"
                      style="{w:pPr/w:pStyle/@w:val}">
                </link>
            </xsl:for-each>
        </links>
    </xsl:template>
    <xsl:template name="link_text">
        <xsl:param name="link"/>
        <xsl:value-of select="$link/w:r/w:t"/>
        <!--
        <xsl:for-each select="$link/w:r/w:t">
            <xsl:value-of select="."/> &#160;-
        </xsl:for-each>
        -->
    </xsl:template>

    <!-- Basic Content -->
    <xsl:template match="w:p">
        <xsl:variable name="style" select="w:pPr/w:pStyle/@w:val"/>
        <xsl:choose>
            <xsl:when test="w:pPr/w:widowControl"></xsl:when>
            <xsl:when test="w:pPr[w:pStyle/@w:val[ starts-with( ., 'ContentsHeading' ) ] ]">
                <xsl:apply-templates select="self::*" mode="toc"/>
            </xsl:when>
            <xsl:when test="w:pPr[w:pStyle/@w:val[ starts-with( ., 'Contents1' ) ] ]"/>
            <xsl:when test="w:pPr[w:pStyle/@w:val[ starts-with( ., 'Heading' ) ] ]">
                <item
                        type="heading"
                        style="{w:pPr/w:pStyle/@w:val}"
                        size="{/w:document/w:styles/w:style[@w:styleId=$style]/w:rPr/w:sz/@w:val}"
                >
                    <content>
                        <xsl:apply-templates/>
                    </content>
                </item>
            </xsl:when>
            <xsl:when test="w:pPr/w:numPr">
                <xsl:apply-templates select="self::*" mode="ol"/>
            </xsl:when>
            <xsl:when test="w:r/w:drawing">
                <item
                        type="image"
                        style="{w:pPr/w:pStyle/@w:val}"
                        size="{/w:document/w:styles/w:style[@w:styleId=$style]/w:rPr/w:sz/@w:val}">
                    <content>
                        <xsl:apply-templates/>
                    </content>
                </item>
            </xsl:when>
            <xsl:otherwise>
                <item
                        type="section"
                        style="{w:pPr/w:pStyle/@w:val}"
                        size="{/w:document/w:styles/w:style[@w:styleId=$style]/w:rPr/w:sz/@w:val}">
                    <content>
                        <xsl:apply-templates/>
                    </content>
                </item>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="w:r">
        <xsl:choose>
            <xsl:when test="w:rPr/w:b[not(@w:val)]"><b><xsl:apply-templates/></b></xsl:when>
            <xsl:when test="w:rPr/w:b[@w:val='true']"><b><xsl:apply-templates/></b></xsl:when>
            <xsl:when test="w:rPr/w:i[not(@w:val)]"><i><xsl:apply-templates/></i></xsl:when>
            <xsl:when test="w:rPr/w:i[@w:val='true']"><i><xsl:apply-templates/></i></xsl:when>
            <xsl:when test="w:rPr/w:highlight"><highlight color="{w:rPr/w:highlight/@w:val}"><xsl:apply-templates/></highlight></xsl:when>
            <xsl:otherwise><xsl:apply-templates/></xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="w:t"><xsl:value-of select="translate(.,'�Â','')"/></xsl:template>

    <!--lists-->
    <!-- next template's only purpose is to fix in a bug in Word where headings
    sometimes get put into List Elements for no reason -->
    <xsl:template match="w:p[w:pPr/w:numPr][starts-with(w:pPr/w:pStyle/@w:val,'Heading')]" mode="ol">
        <xsl:variable name="style" select="w:pPr/w:pStyle/@w:val"/>
        <item
                type='heading'
                style="{w:pPr/w:pStyle/@w:val}"
                size="{/w:document/w:styles/w:style[@w:styleId=$style]/w:rPr/w:sz/@w:val}">
            <content>
                <xsl:apply-templates/>
            </content>
        </item>
    </xsl:template>

    <!-- OK, let's get started. First remove all list items from output -->
    <xsl:template match="w:p[w:pPr/w:numPr]" priority="1" mode="ol"></xsl:template>

    <!--

      Now just match any list item not preceded by a list item.

      We need to match any element not preceded by a listitem *or*
      preceded by a list item at a different level.

      We also need to check if the previous element is a list item and also a Heading, due to
      bug in Word that causes Headers to sometimes get categorized as list elements.

      The preceding element is not a list element:
      not(preceding-sibling::*[1][self::w:p[w:pPr/w:numPr]]) or

      The preceding element is a list element but it is at a different level
      not(preceding-sibling::*[1]/w:pPr/w:numPr/w:ilvl/@w:val = self::w:p/w:pPr/w:numPr/w:ilvl/@w:val)

    -->
    <xsl:template
            match="w:p[w:pPr/w:numPr][
                     (
                       not(preceding-sibling::*[1][self::w:p[w:pPr/w:numPr]]) or
                       not(preceding-sibling::*[1]/w:pPr/w:numPr/w:ilvl/@w:val &lt;= self::w:p/w:pPr/w:numPr/w:ilvl/@w:val)
                     ) and
                     not(starts-with(w:pPr/w:pStyle/@w:val,'Heading'))
                   ]"
            priority="2" mode="ol">
        <xsl:variable name="numId" select="w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="ilvl" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
        <xsl:variable name="listType">
            <xsl:value-of
                    select="/w:document/w:numbering/w:abstractNum[@w:abstractNumId=$numId]/w:lvl[@w:ilvl=$ilvl]/w:numFmt/@w:val"/>
        </xsl:variable>
        <xsl:variable name="elTypeAttr">
            <xsl:call-template name="getElTypeAttr">
                <xsl:with-param name="listType" select="$listType"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="elName">
            <xsl:choose>
                <xsl:when test="$listType='bullet'">ul</xsl:when>
                <xsl:otherwise>ol</xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <item heading="false" type="list">
            <content>
                <xsl:element name="{$elName}">
                    <xsl:attribute name="type">
                        <xsl:value-of select="$elTypeAttr"/>
                    </xsl:attribute>
                    <xsl:attribute name="wtype">
                        <xsl:value-of select="$listType"/>
                    </xsl:attribute>
                    <xsl:apply-templates select="." mode="ordered-list"/>
                </xsl:element>
            </content>
        </item>
    </xsl:template>

    <!-- now match the list item and its next sibling, recursively so you get the entire group -->
    <xsl:template match="w:p[w:pPr/w:numPr]" mode="ordered-list">
        <xsl:variable name="level" select="w:pPr/w:numPr/w:ilvl/@w:val"/><!-- item level -->
        <xsl:variable name="numId" select="w:pPr/w:numPr/w:numId/@w:val"/><!-- id of of list style -->
        <xsl:variable name="indent">
            <xsl:choose>
                <xsl:when test="w:pPr/w:ind/@w:left">
                    <xsl:value-of select="w:pPr/w:ind/@w:left"></xsl:value-of>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
        </xsl:variable><!-- indent of of list style -->
        <li lvl="{$level}">
            <xsl:apply-templates/>
        </li>

        <!-- This block is a duplicate of block in previous template.. TODO: avoid duplication? -->
        <xsl:variable name="listType">
            <xsl:value-of
                    select="/w:document/w:numbering/w:abstractNum[@w:abstractNumId=$numId]/w:lvl[@w:ilvl=$level]/w:numFmt/@w:val"/>
        </xsl:variable>
        <xsl:variable name="elTypeAttr">
            <xsl:call-template name="getElTypeAttr">
                <xsl:with-param name="listType" select="$listType"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="elName">
            <xsl:choose>
                <xsl:when test="$listType='bullet'">ul</xsl:when>
                <xsl:otherwise>ol</xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- end duplicate block -->

        <xsl:choose>
            <!-- case of node is a different ilvl (indent level) than preceding node -->
            <xsl:when test="following-sibling::*[1][self::w:p[w:pPr/w:numPr/w:ilvl/@w:val!=$level]]">
                <xsl:if test="following-sibling::*[self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]][1]"><!-- avoid orphan list el -->

                    <!-- get the attributes for the bullet styles -->
                    <xsl:variable name="nextListType">
                        <xsl:call-template name="getListType">
                            <xsl:with-param name="nextItem"
                                            select="following-sibling::*[self::w:p[w:pPr/w:numPr/w:numId/@w:val!=$level]][1]"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:variable name="nextElTypeAttr">
                        <xsl:call-template name="getElTypeAttr">
                            <xsl:with-param name="listType" select="$nextListType"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <!-- build the enclosing element (OL or UL) -->
                    <xsl:element name="{$elName}">
                        <xsl:attribute name="type">
                            <xsl:value-of select="$nextElTypeAttr"/>
                        </xsl:attribute>
                        <xsl:attribute name="wtype">
                            <xsl:value-of select="$nextListType"/>
                        </xsl:attribute>
                        <xsl:apply-templates select="
                       following-sibling::*[1]
                         [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val>$level]]
                         [self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]]
                       " mode="ordered-list"/>
                    </xsl:element>
                    <xsl:apply-templates select="
                    following-sibling::*
                      [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]]
                       [self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]]
                     [1]
                    " mode="ordered-list"/>
                </xsl:if>
            </xsl:when>

            <!-- case of node is a different numId (style) than preceding node -->
            <xsl:when test="following-sibling::*[1][self::w:p[w:pPr/w:numPr/w:numId/@w:val!=$numId]]">
                <xsl:if test="following-sibling::*[self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]][1]"><!-- avoid orphan list el -->
                    <!-- get the attributes for the bullet styles -->
                    <xsl:variable name="nextListType">
                        <xsl:call-template name="getListType">
                            <xsl:with-param name="nextItem"
                                            select="following-sibling::*[self::w:p[w:pPr/w:numPr/w:numId/@w:val!=$numId]][1]"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <xsl:variable name="nextElTypeAttr">
                        <xsl:call-template name="getElTypeAttr">
                            <xsl:with-param name="listType" select="$nextListType"/>
                        </xsl:call-template>
                    </xsl:variable>
                    <!-- build the enclosing element (OL or UL) -->
                    <xsl:element name="{$elName}">
                        <xsl:attribute name="type">
                            <xsl:value-of select="$nextElTypeAttr"/>
                        </xsl:attribute>
                        <xsl:attribute name="wtype">
                            <xsl:value-of select="$nextListType"/>
                        </xsl:attribute>
                        <xsl:apply-templates select="following-sibling::*[1]
                          [self::w:p[w:pPr/w:numPr/w:numId/@w:val>$numId]]
                          [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]]
                          " mode="ordered-list"/>
                    </xsl:element>
                </xsl:if>
                <xsl:apply-templates
                        select="following-sibling::*
                          [self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]]
                          [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]]
                          [1]"
                        mode="ordered-list"/>
            </xsl:when>

            <!-- case of node w:ind (indent, e.g. bullet) greater than preceding node -->
            <xsl:when test="following-sibling::*[1][self::w:p[w:pPr/w:ind/@w:left!=$indent]]">
                <!-- get the attributes for the bullet styles -->
                <xsl:variable name="nextListType">
                    <xsl:call-template name="getListType">
                        <xsl:with-param name="nextItem"
                                        select="following-sibling::*[self::w:p[w:pPr/w:numPr/w:numId/@w:val!=$indent]][1]"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:variable name="nextElTypeAttr">
                    <xsl:call-template name="getElTypeAttr">
                        <xsl:with-param name="listType" select="$nextListType"/>
                    </xsl:call-template>
                </xsl:variable>
                <!-- build the enclosing element (OL or UL) -->
                <xsl:element name="{$elName}">
                    <xsl:attribute name="type">
                        <xsl:value-of select="$nextElTypeAttr"/>
                    </xsl:attribute>
                    <xsl:attribute name="wtype">
                        <xsl:value-of select="$nextListType"/>
                    </xsl:attribute>
                    <xsl:apply-templates select="following-sibling::*[1][self::w:p[w:pPr/w:ind/@w:left>$indent]]"
                                         mode="ordered-list"/>
                </xsl:element>
                <xsl:apply-templates select="
                  following-sibling::*
                    [self::w:p[w:pPr/w:ind/@w:left!=$indent]]
                    [self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]]
                    [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]]
                    [1]
                  " mode="ordered-list"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:apply-templates select="
                  following-sibling::*[1]
                    [self::w:p[w:pPr/w:numPr]]
                    [self::w:p[w:pPr/w:numPr/w:numId/@w:val=$numId]]
                    [self::w:p[w:pPr/w:numPr/w:ilvl/@w:val=$level]]
                  " mode="ordered-list"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="getElTypeAttr">
        <xsl:param name="listType"/>
        <xsl:choose>
            <xsl:when test="$listType = 'upperRoman'">I</xsl:when>
            <xsl:when test="$listType = 'lowerRoman'">i</xsl:when>
            <xsl:when test="$listType = 'upperLetter'">A</xsl:when>
            <xsl:when test="$listType = 'lowerLetter'">a</xsl:when>
            <xsl:when test="$listType = 'decimal'">1</xsl:when>
            <xsl:when test="$listType = 'bullet'">0</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="getListType">
        <xsl:param name="nextItem"/>
        <xsl:variable name="nextNumId" select="$nextItem/w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="nextLevel" select="$nextItem/w:pPr/w:numPr/w:ilvl/@w:val"/>
        <xsl:value-of
                select="/w:document/w:numbering/w:abstractNum[@w:abstractNumId=$nextNumId]/w:lvl[@w:ilvl=$nextLevel]/w:numFmt/@w:val"/>
    </xsl:template>
    <!-- End Lists -->

    <!-- images -->
    <xsl:template match="w:drawing">
        <xsl:apply-templates select=".//a:blip"/>
    </xsl:template>
    <xsl:template match="a:blip">
        <xsl:variable name="id" select="@r:embed"/><img>
            <xsl:attribute name="src">
                <xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target"/>
            </xsl:attribute>
            <xsl:attribute name="width">
                <xsl:value-of select="round( ancestor::w:drawing[1]//wp:extent/@cx div 9525 )"/>
            </xsl:attribute>
            <xsl:attribute name="height">
                <xsl:value-of select="round( ancestor::w:drawing[1]//wp:extent/@cy div 9525 )"/>
            </xsl:attribute>
        </img></xsl:template>

    <!-- Links -->
    <xsl:template match="w:hyperlink"><xsl:variable name="id" select="@r:id"/><a><xsl:attribute name="href">
                <xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target"/>
            </xsl:attribute><xsl:apply-templates/></a></xsl:template>

    <!-- tables -->
    <xsl:template match="w:tbl">
        <item type="table" heading="false" style="table">
            <content>
                <table>
                    <xsl:apply-templates/>
                </table>
            </content>
        </item>
    </xsl:template>
    <xsl:template match="w:tr">
        <tr>
            <xsl:apply-templates/>
        </tr>
    </xsl:template>
    <xsl:template match="w:tc">
        <td>
            <xsl:value-of select="."/>
        </td>
    </xsl:template>

    <!-- skip contents of these fields -->
    <xsl:template match="rels:Relationships"/>
    <xsl:template match="text()"/>

</xsl:stylesheet>
