<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output
	  method="html"
	  encoding="UTF-8"
	  indent="yes"
	  doctype-public="-//W3C//Dtd HTML 4.01//EN"
	  doctype-system="http://www.w3.org/tr/html4/strict.dtd"/>

	<xsl:template match="/">
		<html>
			<head>
				<style type="text/css">
					table {
					width: 100%;
					color: #212424;
					margin: 0 0 1em 0;
					font: 80%/150% "Lucida Grande", "Lucida Sans Unicode", "Lucida Sans", Lucida, Helvetica, sans-serif;
					}
					table, tr, th, td {
					margin: 0;
					padding: 0;
					border-spacing: 0;
					border-collapse: collapse;
					}

					/* HEADER */

					thead {
					background: #524123;
					}
					thead tr th {
					padding: 1em 0;
					text-align: center;
					color: #FAF7D4;
					border-bottom: 3px solid #A5D768;
					}

					/* BODY */

					tbody tr td {
					background: #DDF0BD no-repeat top left;
					}
					tbody tr.odd td {
					background-color: #D0EBA6;
					}
					tbody tr th, tbody tr td {
					padding: 0.1em 0.4em;
					border: 1px solid #a6ce39;
					}
					tbody tr th {
					padding-right: 1em;
					text-align: right;
					font-weight: normal;
					background: #c5e894 no-repeat top left;
					text-transform: uppercase;
					}
					td.error {
					background-color: #FF3300;
					color: #FFFFFF;
					}
					td.warning{
					background-color: #FFFF00;
					}
					}
					span.error {
					background-color: #FF3300;
					color: #FFFFFF;
					}
					span.warning {
					background-color: #FFFF00;
					}
				</style>
				
				<title><xsl:value-of select="//ObjectName" /></title>
			</head>

			<body>
				<xsl:apply-templates />
			</body>
		</html>
	</xsl:template>

	<xsl:template name="split">
		<xsl:param name="list" />
		<xsl:param name="delimiter" />
		<xsl:param name="linenum" />
		<xsl:variable name="newlist">
			<xsl:choose>
				<xsl:when test="contains($list, $delimiter)">
					<xsl:value-of select="$list" />
				</xsl:when>

				<xsl:otherwise>
					<xsl:value-of select="concat($list, $delimiter)"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="first" select="substring-before($newlist, $delimiter)" />
		<xsl:variable name="remaining" select="substring-after($newlist, $delimiter)" />
		<tr>
			<th>
				<a>
					<xsl:attribute name="name">
						<xsl:value-of select="$linenum" />
					</xsl:attribute>
					<xsl:value-of select="$linenum" />
				</a>
			</th>
			<td>
				<xsl:if test="count(/errs/err[linenum=$linenum and type='Warning'])>0">
					<xsl:attribute name="class">warning</xsl:attribute>
				</xsl:if>
				<xsl:if test="count(/errs/err[linenum=$linenum and type='Error'])>0">
					<xsl:attribute name="class">error</xsl:attribute>
				</xsl:if>
				<xsl:value-of select="translate($first, ' ',  '&#160;')"/>
			</td>
		</tr>

		<xsl:if test="$remaining">
			<xsl:call-template name="split">
				<xsl:with-param name="list" select="$remaining" />
				<xsl:with-param name="delimiter">
					<xsl:value-of select="$delimiter"/>
				</xsl:with-param>
				<xsl:with-param name="linenum" select="$linenum + 1" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<xsl:template match="errs">
		<table>
			<thead>
				<tr>
					<th colspan="2">USACE GIS Metadata Editor - Validation Results - <xsl:value-of select="//ObjectName" /></th>
				</tr>
			</thead>

			<xsl:choose>
				<xsl:when test="count(err)=0">
					<tr>
						<td>No errors or warnings!</td>
					</tr>
				</xsl:when>
				<xsl:otherwise>
					<tr>
						<th>Summary</th>
						<td>
							<span class="error"><xsl:value-of select="count(err[type='Error'])"/> error(s)</span>,
							<span class="warning"><xsl:value-of select="count(err[type='Warning'])"/> warning(s)</span>.
						</td>
					</tr>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:apply-templates />
		</table>
	</xsl:template>

	<xsl:template match="err">
		<tr>
			<td>&#160;</td>
		</tr>
		<xsl:apply-templates />
	</xsl:template>

	<xsl:template match="type[.='Error']">
		<tr>
			<th>Type</th>
			<td>
				<span class="error">
					<xsl:apply-templates />
				</span>
			</td>
		</tr>
	</xsl:template>

	<xsl:template match="type[.='Warning']">
		<tr>
			<th>Type</th>
			<td>
				<span class="warning">
					<xsl:apply-templates />
				</span>
			</td>
		</tr>
	</xsl:template>

	<xsl:template match="linenum">
		<tr>
			<th>Line No</th>
			<td>
				<a>
					<xsl:attribute name="href">
						#<xsl:value-of select="." />
					</xsl:attribute>
					<xsl:value-of select="." />
				</a>
			</td>
		</tr>
	</xsl:template>

	<xsl:template match="errid">
		<tr>
			<th>Path</th>
			<td>
				<xsl:apply-templates />
			</td>
		</tr>
	</xsl:template>

	<xsl:template match="message">
		<tr>
			<th>Description</th>
			<td>
				<xsl:apply-templates />
			</td>
		</tr>
	</xsl:template>

	<xsl:template match="src">
		<table>
			<thead>
				<tr>
					<th colspan="2">Metadata record</th>
				</tr>
			</thead>
			<xsl:call-template name="split">
				<xsl:with-param name="list">
					<xsl:value-of select="."/>
				</xsl:with-param>
				<xsl:with-param name="delimiter">&#xD;&#xA;</xsl:with-param>
				<xsl:with-param name="linenum">1</xsl:with-param>
			</xsl:call-template>
		</table>
	</xsl:template>

</xsl:stylesheet>
