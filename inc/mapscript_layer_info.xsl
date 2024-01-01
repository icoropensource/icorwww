<?xml version="1.0" encoding="windows-1250" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

<xsl:template match="LAYERS">
   <html>
   <head>
   <title>Informacje o warstwach mapy</title>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1250" />
   <meta name="pragma" content="no-cache" />
   <link rel="STYLESHEET" type="text/css" href="/icormanager/icor.css" />
   </head>
   <body>
      <table cellpadding="4" cellspacing="0" border="0">
         <xsl:apply-templates select="LAYER" />
      </table>
   </body>
   </html>  
</xsl:template>

<xsl:template match="LAYER">
   <tr class="objectseditrow">
       <td colspan="3">
          <font size="4"><u><b><img src="/icormanager/images/wfxtree/items/Add_Command_InsertMap.png" />&#160;<xsl:value-of select="@name" disable-output-escaping="yes" /></b></u></font>&#160;
       </td>
   </tr>           
   <xsl:apply-templates select="FIELD" />
   <tr>
      <td colspan="3">
         <xsl:if test="position()!=last()">
            <hr/>
         </xsl:if>
      </td>
   </tr>
</xsl:template>

<xsl:template match="FIELD">
   <tr class="objectseditrow">
      <td width="30px">
         &#160;
      </td>   
      <td align="right" nowrap="true">   
         <b><xsl:value-of select="@name" disable-output-escaping="yes" />:</b>&#160;
      </td>
      <td align="left">
         <xsl:value-of select="@value" disable-output-escaping="yes" />&#160;
      </td>       
   </tr>    
</xsl:template>
</xsl:stylesheet>
