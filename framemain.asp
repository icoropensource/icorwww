<%@ codepage=65001 LANGUAGE="VBSCRIPT" %><!--#include file ="definc_createIISI.asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="window-target" content="_top" />
<title>ICOR (<%=Dession("UserName")%>)</title>
<script language="javascript">
if (window!=top) top.location.href=location.href;
</script>
<style>
   html {background-color:<%=Dession("FRAME_BGCOLOR")%>;font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;}
   frameset {margin-left:0px;margin-right:0px;margin-bottom:0px;margin-top:0px;}
   #tocpane {border:solid 1px <%=Dession("FRAME_TOC_BGCOLOR")%>;}
   #text {border:solid 1px <%=Dession("FRAME_TEXT_BGCOLOR")%>;}
</style>
</head>
<FRAMESET frameSpacing="0" border="NO" frameBorder="0"  rows="55,*">
 <FRAME name="NAVBAR" src="navbar.asp" marginWidth="0" marginHeight="0" scrolling="no" frameBorder="0">
   <FRAMESET frameSpacing="12" name="framesetTOCPane" border="1" cols="20%,*" frameBorder="1">
     <FRAME id="tocpane" name="TOCPANE" src="contents.asp" marginWidth="0" marginHeight="0" border="1" frameBorder="1" borderColor=<%=Dession("FRAME_BGCOLOR")%>>
     <FRAME id="text" name="TEXT" src="startpage.asp" target="TEXT" marginWidth="0" marginHeight="0" border="1" frameBorder="1" borderColor=<%=Dession("FRAME_BGCOLOR")%>>
   </FRAMESET>
</FRAMESET>
</html>