<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:t="urn:schemas-microsoft-com:time" xmlns="http://www.w3.org/TR/REC-html40" xmlns:tool>
<head><link rel=STYLESHEET type="text/css" href="icor.css" title="SOI">
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="pl">
<meta name="description" content="Unload processors">
<META NAME="Pragma" CONTENT="no-cache">
<META NAME="Cache-Content" CONTENT="no-cache">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Content" CONTENT="no-cache">
<META HTTP-EQUIV="Expires" CONTENT="-1">
<meta name="keywords" content="ICOR object oriented database WWW information management repository">
<meta name="author" content="<%=Application("META_AUTHOR")%>">
<?IMPORT namespace="t" implementation="#default#time2">
<title>Unload processors</title>
<base target="TEXT">
</head>
<body topmargin="10">
<!--#include file ="definc_createIISI.asp"-->
<%

call RefreshProcessor()
response.write "<h1>Process pool refreshed.</h1>"

%>
</body></html>
