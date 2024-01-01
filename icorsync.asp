<!--#include file ="definc_createIISI.asp"--><%
Response.Charset = "utf-8"
Response.CacheControl = "no-cache"
Response.ExpiresAbsolute = #1/1/1999 1:50:00 AM#
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.ContentType = "text/xml"
if request.querystring("state")<>"" then
   if request.querystring("mode")="get" then
      astate=CLng(request.querystring("state"))
      avalue=ICORSyncGetState(astate)
      aitems=Split(avalue,SPLIT_CHAR_PARAM)
      aname=aitems(lbound(aitems))
      avalue=aitems(ubound(aitems))
      response.write "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbcrlf
      response.write "<DATA>" & vbcrlf
      response.write "<STATE name=""" & aname & """ value=""" & avalue & """ />" & vbcrlf
      response.write "</DATA>" & vbcrlf
      if (avalue="BAD") or (avalue="OK") then
         ret=ICORSyncDelState(astate)
      end if
   end if
end if
%><!--#include file ="definc_disposeIISI.asp"-->
