<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:html="http://www.w3.org/TR/REC-html40">
<xsl:output method="html" encoding="utf-8" indent="yes" />

<xsl:template match="/">
<!--<xsl:text disable-output-escaping='yes'>&lt;!DOCTYPE html&gt;&#10;</xsl:text>-->
  <html>
    <head>
      <style>
        body {
          font-family: Verdana, Arial, sans-serif;
        }
        table, th, td { 
          border: 0.5pt solid;
          border-color: gainsboro;
          border-collapse: collapse;
          padding: 4px;
        }
        th { 
          background-color: Thistle;
          font-size: 0.9em;
          <!-- cursor: pointer; -->
        }
        td {
          vertical-align: top;
          font-size: 0.75em;
        }
		li {
		  font-size: 0.7em;
		}
      </style>

      <script>
      </script>
    </head>

    <body>
      <h1>System Controller Report</h1>
	  
	  <ul>
	  <xsl:for-each select="/ss:Workbook/ss:Worksheet">
		<li><a>
		  <xsl:attribute name="href">
			<xsl:value-of select="concat('#',@ss:Name)" />
		  </xsl:attribute>
		  <xsl:value-of select="@ss:Name"/>
		</a></li>
	  </xsl:for-each>
      </ul>  

	  <xsl:for-each select="/ss:Workbook/ss:Worksheet">
	  
		  <!-- <h2 id="<xsl:value-of select="@ss:Name"/>"><xsl:value-of select="@ss:Name"/></h2> -->
		  <h2>
            <xsl:attribute name="id">
			  <xsl:value-of select="@ss:Name" />
			</xsl:attribute>
			<xsl:value-of select="@ss:Name"/>
		  </h2>

		  <table border="1">
			<xsl:for-each select="ss:Table/ss:Row">
			  <tr>
				<xsl:for-each select="ss:Cell">
				  <td><xsl:value-of select="ss:Data"/></td>
				</xsl:for-each>
			  </tr>
			</xsl:for-each>
		  </table>
	  
	  </xsl:for-each> 
	  
    </body>
  </html>
</xsl:template>

</xsl:stylesheet>