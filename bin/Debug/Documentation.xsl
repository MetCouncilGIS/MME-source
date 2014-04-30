<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:template match="doc">
		<xsl:apply-templates select="./assembly/name" />
		<H2>
			Types
		</H2>
		<DL>
			<xsl:apply-templates select="./members/member[starts-with(@name,'T:')]" />
		</DL>
	</xsl:template>
	<xsl:template match="name">
		<TABLE width="100%" border="1" cellspacing="0" cellpadding="10" bgcolor="lightskyblue">
			<TR>
				<TD align="center">
					<H1>
						<xsl:text>Assembly: </xsl:text>
						<xsl:value-of select="." />
					</H1>
				</TD>
			</TR>
		</TABLE>
	</xsl:template>
	<xsl:template match="member[starts-with(@name,'T:')]">
		<DT>
			<A>
				<xsl:attribute name="href">GiveTypeMemberListHelp.aspx?Type=<xsl:value-of select="substring-after(@name,':')" /></xsl:attribute>
				<B>
					<xsl:value-of select="substring-after(@name,'.')" />
				</B>
			</A>
		</DT>
		<DD>
			<xsl:apply-templates select="remarks | summary" />
		</DD>
		<P></P>
	</xsl:template>
	<xsl:template match="remarks | summary">
		<xsl:apply-templates />
	</xsl:template>
	<xsl:template match="c">
		<CODE>
			<xsl:apply-templates />
		</CODE>
	</xsl:template>
	<xsl:template match="para">
		<P>
			<xsl:apply-templates />
		</P>
	</xsl:template>
	<xsl:template match="paramref">
		<I>
			<xsl:apply-templates />
		</I>
	</xsl:template>
	<xsl:template match="see">
		<A>
			<xsl:variable name="MemberString" select="substring-after(@cref,':') " />
			<xsl:variable name="TypeString" select="concat(concat( substring-before($MemberString,'.'),  '.'), substring-before(substring-after($MemberString,'.'),'.'))" />
			<xsl:choose>
				<xsl:when test="starts-with(@cref,'T:')">
					<xsl:attribute name="href">
					GiveTypeHelp.aspx?Type=
					<xsl:value-of select="$TypeString" />
				</xsl:attribute>
					<xsl:value-of select="substring-after(substring-after($MemberString,'.'),'.')" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="href">
					GiveTypeMemberHelp.aspx?Type=
					<xsl:value-of select="$TypeString" />
					&amp;Member=
					<xsl:value-of select="$MemberString" />
				</xsl:attribute>
					<xsl:value-of select="substring-after(substring-after($MemberString,'.'),'.')" />
				</xsl:otherwise>
			</xsl:choose>
			<xsl:apply-templates />
		</A>
	</xsl:template>
	<xsl:template match="list[@type='table']">
		<TABLE border="1">
			<xsl:apply-templates mode="table" />
		</TABLE>
	</xsl:template>
	<xsl:template match="list[@type='bullet']">
		<UL>
			<xsl:apply-templates mode="list" />
		</UL>
	</xsl:template>
	<xsl:template match="list[@type='number']">
		<OL>
			<xsl:apply-templates mode="list" />
		</OL>
	</xsl:template>
	<xsl:template match="listheader" mode="table">
		<THEAD>
			<xsl:apply-templates mode="table" />
		</THEAD>
	</xsl:template>
	<xsl:template match="listheader" mode="list">
		<LI>
			<B>
				<xsl:apply-templates mode="list" />
			</B>
		</LI>
	</xsl:template>
	<xsl:template match="item" mode="table">
		<TR>
			<xsl:choose>
				<xsl:when test="count(child::*) = 0">
					<xsl:choose>
						<xsl:when test="name(parent::*)='listheader'">
							<TH>
								<xsl:apply-templates mode="table" />
							</TH>
						</xsl:when>
						<xsl:otherwise>
							<TD>
								<xsl:apply-templates mode="table" />
							</TD>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:apply-templates mode="table" />
				</xsl:otherwise>
			</xsl:choose>
		</TR>
	</xsl:template>
	<xsl:template match="item" mode="list">
		<LI>
			<xsl:apply-templates mode="list" />
		</LI>
	</xsl:template>
	<xsl:template match="term | description" mode="table">
		<xsl:choose>
			<xsl:when test="name(parent::*)='listheader'">
				<TH>
					<xsl:apply-templates mode="table" />
				</TH>
			</xsl:when>
			<xsl:when test="name(parent::*)='item'">
				<TD>
					<xsl:apply-templates mode="table" />
				</TD>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="term" mode="list">
		<xsl:apply-templates mode="list" />
		<xsl:text> - </xsl:text>
	</xsl:template>
	<xsl:template match="description" mode="list">
		<xsl:apply-templates mode="description" />
	</xsl:template>
	<xsl:template match="code">
		<pre>
			<xsl:apply-templates />
		</pre>
	</xsl:template>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:apply-templates />
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>
