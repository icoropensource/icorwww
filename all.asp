<%
'Session.CodePage = 65001
Response.CodePage = 65001
Response.CharSet = "utf-8"

Sub NoCache()
  Response.Buffer = True
  Response.Expires = 0
  Response.AddHeader "Expires",0
  Response.AddHeader "Pragma","no-cache"
  Response.AddHeader "Cache-Control","no-cache, private, post-check=0, pre-check=0, max-age=0"
  Response.ExpiresAbsolute = Now() - 1
  Response.CacheControl = "no-cache"
End Sub
%><!--#include file="inc/forms/validation.asp"--><!--#include file="inc/page/db.asp"--><!--#include file ="inc/dbsession.asp"--><!--#include file="inc/page/security.asp"--><!--#include file="inc/forms/json.asp"--><%
set db = new DBConnection
%>