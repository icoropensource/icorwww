<%@ CodePage=65001 LANGUAGE="VBSCRIPT" %><%
sorigin=Request.ServerVariables("HTTP_ORIGIN")
if sorigin="" then
   sorigin="*"
end if
Response.AddHeader "Access-Control-Allow-Origin", sorigin
if (Request.ServerVariables("REQUEST_METHOD")<>"GET") and (Request.ServerVariables("REQUEST_METHOD")<>"POST") then
   Response.End
end if
%><!-- #include file="all.asp" --><%
Call NoCache

amode=ReplaceIllegalCharsLev0Text(mid(Request.QueryString("m"),1,30))
bmode=ReplaceIllegalCharsLev0Text(mid(Request.Form("m"),1,30))
if amode="" and bmode<>"" then
   amode=bmode
end if

if amode="SessionEnd" then
   astatus="0"
   Dession("uid") = -1
   Dession("UserName") = ""
   Dession("CMS_UserID")=""
   Dession("CMS_IdentityID")=""
   response.write "{""status"":" & astatus & "}"
end if

if amode="SessionCheck" then
   astatus=1
   ainfo="użytkownik nie jest zalogowany"
   aCMS_UserID=Dession("CMS_UserID")
   aCMS_IdentityID=Dession("CMS_IdentityID")
   aUID=CStr(Dession("uid"))
   if (aCMS_UserID<>"") or (aCMS_IdentityID<>"") or ((aUID<>"") and (aUID<>"-1")) then
      astatus="0"
      ainfo=""
   end if
   response.write "{""status"":" & astatus & ",""info"":""" & ainfo & """}"
end if
%>