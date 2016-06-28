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
                exclude-result-prefixes="*">

    <xsl:output method="xml" indent="yes"
                encoding="UTF-8"
            />

    <xsl:template match="/">
    doc = {
       <xsl:apply-templates />
    }

    </xsl:template>

    <!-- document properties .e.g title -->
    <xsl:template match="head">
       properties : {
        <xsl:apply-templates/>
        },
    </xsl:template>
    <xsl:template match="head/*"><xsl:value-of select="local-name()"/>: '<xsl:value-of select="."/>',</xsl:template>

    <xsl:template match="body">
        items : [
        <xsl:apply-templates/>
        ],
    </xsl:template>
    <xsl:template match="toc">
        {
          type : 'toc',
          items : [
        <xsl:apply-templates/>
        ]
        },
    </xsl:template>
    <xsl:template match="item">{
          type : '<xsl:value-of select="@type"/>',
          style : '<xsl:value-of select="@style"/>',
          content : '<xsl:copy-of select="content"/>'
        },</xsl:template>

    <xsl:template match="text()"><xsl:value-of select="translate(.,'�','')"/></xsl:template>

</xsl:stylesheet>
