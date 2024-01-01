<!--#include file ="inc/dbsession.asp"--><%
Dession("uid") = -1
Dession("UserName") = ""
Dession("relogin2")=0
Dession("TmpServerFileCounter")=0
Dession("TmpClientFileCounter")=1
Dession("LastVisitHistoryID")=1
Dession("logout")=1
Dession("CMS_UserID")=""
Dession("CMS_IdentityID")=""
'call Session.Abandon
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="pl">
<meta name="description" content="Logout">
<META NAME="Pragma" CONTENT="no-cache">
<META NAME="Cache-Content" CONTENT="no-cache">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Content" CONTENT="no-cache">
<META HTTP-EQUIV="Expires" CONTENT="-1">
<meta name="keywords" content="ICOR object oriented database WWW information management repository">
<meta name="author" content="<%=Application("META_AUTHOR")%>">
<title>Logout...</title>
<base target="TEXT">
</head>
<body>
<script language="javascript">
window.top.close();
</script>
</body></html>
