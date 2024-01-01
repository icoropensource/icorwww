<!--#include file ="inc/dbsession.asp"--><!--#include file ="inc/cache.asp"--><!--#include file ="definc_ICORAPI2.asp"--><%
SPLIT_CHAR_PARAM="$*$!"
SPLIT_CHAR_VALUE="$*!$"
IGNORE_VISIT_HISTORY=0

Function ICORExecuteMethod(amethod,afieldname,aoid,avalue,adefault)
   Dim aicor,apos,aclass,bmethod,acid,anostringvalue
   ICORExecuteMethod=adefault
   amethod=replace(amethod,"_","\")
   amethod=replace(amethod,"/","\")
   apos=InStrRev(amethod,"\")
   if apos>0 then
      aclass=Mid(amethod,1,apos-1)
      bmethod=Mid(amethod,apos+1)
      acid=GetClassID(0,aclass)
      if acid>255 then
         anostringvalue=1
         ICORExecuteMethod=CStr(ExecuteMethod(Dession("uid"),CLng(acid),CStr(bmethod),CStr(afieldname),CLng(aoid),CStr(avalue),anostringvalue))
      end if
   end if
End Function

Function ICORExecuteMethodAsUID(auid,amethod,afieldname,aoid,avalue,adefault)
   Dim aicor,apos,aclass,bmethod,acid,anostringvalue
   ICORExecuteMethodAsUID=adefault
   amethod=replace(amethod,"_","\")
   amethod=replace(amethod,"/","\")
   apos=InStrRev(amethod,"\")
   if apos>0 then
      aclass=Mid(amethod,1,apos-1)
      bmethod=Mid(amethod,apos+1)
      acid=GetClassID(0,aclass)
      if acid>255 then
         anostringvalue=1
         ICORExecuteMethodAsUID=CStr(ExecuteMethod(CLng(auid),CLng(acid),CStr(bmethod),CStr(afieldname),CLng(aoid),CStr(avalue),anostringvalue))
      end if
   end if
End Function

function GetXMLHTTPDataPost(aurl,aparams,asecs)
   dim aitertime,xmlHttp
   GetXMLHTTPDataPost=""
   if Response.IsClientConnected then
      set xmlHttp=server.createobject("MSXML2.ServerXMLHTTP")
      xmlHttp.setTimeouts 0, 120*1000, 15*60*1000, 30*60*1000
      xmlHttp.open "POST",aurl,true
      xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
      xmlHttp.send aparams
      aitertime=0
      Do While xmlHttp.ReadyState<>4
         If (not Response.IsClientConnected) or (aitertime>=asecs) Then
            Exit Do
         end if
         xmlHttp.waitForResponse 1
         aitertime=aitertime+1
      loop
      if xmlHttp.ReadyState=4 then
         if xmlHttp.Status=200 then
            GetXMLHTTPDataPost=xmlHttp.responseText
         else
            xmlHttp.Abort
         end if
      else
         xmlHttp.Abort
      end if
      set xmlHttp=nothing
   end if
end function

Function GetHTTPURL(aurl)
   Dim xmlhttp
   GetHTTPURL=""
   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
   xmlhttp.Open "POST", aurl, false
   xmlhttp.Send ""
   set xmlhttp=Nothing
End Function

Sub OutputStr(s)
   'dim ret
   'ret=ICORExecuteMethodAsUID(0,"CLASSES\Library\NetBase\WWW\Server\WWWOutputStr", "", -1, s, "")
End Sub

Sub LogLine(s)
   dim ret
   s=Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & " | " & s)
   ret=GetHTTPURL("http://localhost:9095/logprint/" & s)
End Sub

sub ICORSessionInit()
   'Session("_sessioninit")=0
   asessiontoken=Dession.GetSessionToken()
   if Dession.firstSession then
      Dession("ICORServer") = ""
      Dession("uid") = -1
      Dession("UserName") = ""
      Dession("DefaultSQLRecordsPerPage")="20"
      Dession("relogin2")=0
      Dession("TmpServerFileCounter")=0
      Dession("TmpClientFileCounter")=1
      Dession("logout")=0

      Dession("LastVisitHistoryID")=1
      Dession("UI_SKIN")="redmond"
      Dession("FRAME_BGCOLOR")="#f6f6f6"
      Dession("FRAME_TOC_BGCOLOR")="#cccccc"
      Dession("FRAME_TEXT_BGCOLOR")="#cccccc"
      
      atime=Application("LastProcessPoolRefresh")
      if atime="" then
         atime=Now()
         Application.Lock
         Application("LastProcessPoolRefresh")=atime
         Application.Unlock
      end if
      Dession("LastProcessPoolRefresh")=atime
   end if
end sub

'if Session("_sessioninit")=1 then
'   s=CStr(Request.ServerVariables("HTTP_COOKIE"))
'   if len(s)=0 then
'      Response.Redirect "/icormanager/"
'      Response.Flush
'      Response.End
'   end if
'   call ICORSessionInit
'end if

MAX_HISTORY=30

set cache = new ClsCache
with cache
   .name = "test"
   .interval = "s" 'n-minute
   .intervalValue = 15
end with

Sub RefreshSession()
   if CStr(Dession("logout"))=CStr(1) then
      'Session.Abandon
      Response.Redirect "/icormanager/"
      Response.Flush
      Response.End
   end if
   if Dession("LastProcessPoolRefresh")<>Application("LastProcessPoolRefresh") then
      'For Each Key In Session.Contents
      '  If Not IsObject(Session.Contents(Key)) Then
      '     if (mid(Key,1,5)="PACL_") or (mid(Key,1,4)="SID_") or (mid(Key,1,8)="CHID_TC_") or (mid(Key,1,8)="CHID_ALV_") or (mid(Key,1,8)="SIDCALV_") then
      '        Session(Key)=""
      '     end if
      '  End If
      'Next
      Dession("LastProcessPoolRefresh")=Application("LastProcessPoolRefresh")
   end if
End Sub

Function ExecuteMethodRetS(amethod,afieldname,aoid,avalue,adefault)
   ExecuteMethodRetS=ICORExecuteMethod(amethod,afieldname,aoid,avalue,adefault)
End Function

Function ExecuteMethodRetI(amethod,afieldname,aoid,avalue,adefault)
   dim ret
   ExecuteMethodRetI=ICORExecuteMethod(amethod,afieldname,aoid,avalue,adefault)
End Function

Sub ICORClose()
   Dim aicor
   Call RepositoryChange(0,"Close",-1,-1,"","","")
End Sub

Function ICORExists()
   'On Error Resume Next
   ICORExists = ICORExecuteMethodAsUID(0, "CLASSES\Library\NetBase\WWW\Server\ICORExists", "", -1, -1, -1)
End Function

function GetICORServerStatus()
   GetICORServerStatus=ICORExecuteMethodAsUID(0,"CLASSES\Library\NetBase\WWW\Server\WWWCheckServer", "", -1, "", "")
end function

Function ICORSyncAddState(aname,avalue)
   ICORSyncAddState=ICORExecuteMethod("CLASSES\System\SystemDictionary\Synchronization\State\AddState",aname,-1,avalue,"-1")
   ICORSyncAddState=CLng(ICORSyncAddState)
End Function

Function ICORSyncGetState(astate)
   ICORSyncGetState=ICORExecuteMethod("CLASSES\System\SystemDictionary\Synchronization\State\GetState","",astate,"","")
End Function

Function ICORSyncSetState(astate,aname,avalue)
   ICORSyncSetState=ICORExecuteMethod("CLASSES\System\SystemDictionary\Synchronization\State\SetState",aname,astate,avalue,"-1")
   ICORSyncSetState=CLng(ICORSyncSetState)
End Function

Function ICORSyncDelState(astate)
   ICORSyncDelState=ICORExecuteMethod("CLASSES\System\SystemDictionary\Synchronization\State\DelState","",astate,"","")
   ICORSyncDelState=CLng(ICORSyncDelState)
End Function

Function ICORSyncWriteClientScript(astate)
   Response.Write "<script language='javascript' defer>" & chr(10)
   Response.Write "window.top.frames('NAVBAR').registerStateBadOK(" & CStr(astate) & ");" & chr(10)
   Response.Write "</script>" & chr(10)
   ICORSyncWriteClientScript=astate
End Function

Function GetICORVariable(aname)
   GetICORVariable=GetVariable(Dession("uid"),CStr(aname))
End Function

Function SetICORVariable(aname,avalue)
   SetVariable Dession("uid"),CStr(aname),CStr(avalue)
End Function

Function ConvertMemoToString(amemo, ahighchars)
   Dim s1, i, c
   s1 = ""
   If ahighchars Then
      For i = 1 To Len(amemo)
         c = Mid(amemo, i, 1)
         If (c >= " ") And (c <= Chr(127)) And (c <> """") And (c <> "'") Then
            s1 = s1 + c
         Else
            s1 = s1 + "&#" + CStr(Asc(c)) + ";"
         End If
      Next
   Else
      For i = 1 To Len(amemo)
         c = Mid(amemo, i, 1)
         If (c < " ") Or (c = """") Or (c = "'") Then
            s1 = s1 + "&#" + CStr(Asc(c)) + ";"
         Else
            s1 = s1 + c
         End If
      Next
   End If
   ConvertMemoToString = s1
End Function

Function GetServerVariables()
   GetServerVariables = ""
   For Each key In Request.ServerVariables
      ss = ConvertMemoToString(Request.ServerVariables(key), False)
      GetServerVariables = GetServerVariables & SPLIT_CHAR_PARAM & key & SPLIT_CHAR_VALUE & ss
   Next
End Function

Function ICORLogin()
   Dim s, key, ss
   Dession("UserImpersonated")=0
   s = GetServerVariables()
   ICORLogin = ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\WWWLogin", "", -1, s, "-1")
   if Mid(ICORLogin,1,1)="#" then
      Dession("UserImpersonated")=1
      ICORLogin=Mid(ICORLogin,2,1000)
   end if
   if ICORLogin="" then
      ICORLogin="-1"
   end if
   ICORLogin=CLng(ICORLogin)
End Function

Function ICORCheckLogin(auser,apassword)
   ICORCheckLogin = ExecuteMethodRetI("CLASSES\Library\NetBase\WWW\Server\ICORGetUIDByUserPassword", auser, -1, apassword, "-1")
End Function

Function ICORLoginCheckSession()
   Dim s, key, ss
   s = GetServerVariables()
   ICORLoginCheckSession = ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\WWWLoginCheckSession", "", -1, s, "-1")
End Function

Function GetInputValue(aname)
   if Request.QueryString(aname)<>"" then
      GetInputValue=Request.QueryString(aname)
   else
      GetInputValue=Request.Form(aname)
   end if
End Function

Function Stream_BinaryToString(Binary, CharSet)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeBinary
  BinaryStream.Open
  BinaryStream.Write Binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  If Len(CharSet) > 0 Then
    BinaryStream.CharSet = CharSet
  Else
    BinaryStream.CharSet = "ascii"
  End If
  Stream_BinaryToString = BinaryStream.ReadText
End Function


'********************************************
'Function PadHex(anumber)
'   PadHex=Hex(CLng(anumber))
'   PadHex=String(8-len(PadHex),"0")+PadHex
'End Function

Function Proc_ExecuteMethodRetS(aprocessor,amethod,afieldname,aoid,avalue,adefault)
   Proc_ExecuteMethodRetS=ICORExecuteMethod(amethod,afieldname,aoid,avalue,adefault)
end Function

Function Proc_ExecuteMethodNoWait(aprocessor,amethod,afieldname,aoid,avalue)
   Proc_ExecuteMethodNoWait=ICORExecuteMethod(amethod,afieldname,aoid,avalue,adefault)
end Function

Function getParams(aparams)
   Dim i,aitems,avalues
   Set getParams = Server.CreateObject("Scripting.Dictionary")
   aitems=Split(aparams,SPLIT_CHAR_PARAM)
   for i=lbound(aitems) to ubound(aitems)
      avalues=Split(aitems(i),SPLIT_CHAR_VALUE)
      getParams.Add avalues(lbound(avalues)),avalues(ubound(avalues))
   next
End Function

Function putParams(aparams)
   Dim i,akeys
   putParams=""
   akeys=aparams.Keys
   For i=0 To aparams.Count-1
      putParams=putParams & SPLIT_CHAR_PARAM & akeys(i) & SPLIT_CHAR_VALUE & aparams(akeys(i))
   Next
End Function

Function getParamsAsList(aparams)
   Dim i,aitems,avalues,aret
   aitems=Split(aparams,SPLIT_CHAR_PARAM)
   aret=Array()
   redim aret(ubound(aitems))
   for i=lbound(aitems) to ubound(aitems)
      avalues=Split(aitems(i),SPLIT_CHAR_VALUE)
      aret(i)=avalues
   next
   getParamsAsList=aret
End Function
'********************************************

sub CheckSession
   RefreshSession
   dim suid,auid
   suid=CStr(Dession("uid"))
   if suid<>"" then
      auid=CLng(suid)
   else
      auid=-1
   end if
   If auid>0 Then
      s=Dession("username")
      if (s="") or (s="Administrator") then
         Dession("username")=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\GetUserByUID", "", -1, "", "")
      end if
   end if
   If (auid<0) or (Request.QueryString("relogin")="1") Then
      Dession("uid")=ICORLogin()
      If Dession("uid") > 0 Then
         Dession("username")=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\GetUserByUID", "", -1, "", "")
         Dession("_passwordcheck")=99
      else
         Response.Redirect "/icormanager/"
         Response.Flush
         Response.End
      end if
   End If
end sub

function GetFormValuesAsString()
   Dim Key
   GetFormValuesAsString=""
   For Each Key In Request.Form
      GetFormValuesAsString=GetFormValuesAsString & "<input type=hidden name='" & CStr(Key) & "' value='" & Server.HTMLEncode(Request.Form(Key)) & "'>"
   Next
end function

sub UpdateLastVisitHistoryValue(aname,byval avalue)
   dim aid
   if IGNORE_VISIT_HISTORY=1 then
      Exit Sub
   end if
   aid=Dession("LastVisitHistoryIDForUpdate")
   avalue=trim(avalue)
   if avalue<>"" then
      if len(avalue)>60 then
         avalue=left(avalue,50) & "..."
      end if
      Dession("lastvisithistory_value_" & CStr(aid))=Dession("lastvisithistory_value_" & CStr(aid)) & "<TR VALIGN=top class=objectseditrow><td class=objectseditdatafieldname>" & aname & ":</td><td class=objectseditdatafieldvalue>" & avalue & "</td></tr>"
   end if
end sub

sub RegisterVisitHistory(aname,aurl,aform,amode)
   if IGNORE_VISIT_HISTORY=1 then
      Exit Sub
   end if
   if Dession("lastvisithistory_ignore_click")=1 then
      Dession("lastvisithistory_ignore_click")=0
      Exit Sub
   end if
   aid=Dession("LastVisitHistoryID")
   Dession("LastVisitHistoryIDForUpdate")=aid
   if aurl="" then
      aurl=Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString
   end if
   Dession("lastvisithistory_url_" & CStr(aid))=aurl
   Dession("lastvisithistory_form_" & CStr(aid))=aform
   Dession("lastvisithistory_name_" & CStr(aid))=aname
   Dession("lastvisithistory_time_" & CStr(aid))=CStr(Now)
   Dession("lastvisithistory_mode_" & CStr(aid))=amode
   Dession("lastvisithistory_value_" & CStr(aid))=""
   aid=aid+1
   if aid>MAX_HISTORY then
      aid=1
   end if
   Dession("LastVisitHistoryID")=aid
end sub

sub GetICORSessionVars
   s=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\WWWGetSessionVars", "", -1, "", "")
   if s<>"" then
      aitems=getParamsAsList(s)
      for i=lbound(aitems) to ubound(aitems)
         Dession(aitems(i)(0))=aitems(i)(1)
      next
   end if
end sub

Function Process(paction)
   Dim iaction, bc, errkind, s, i, scs, alen,duid
   dim ado_stream,xml_dom,xml_file1
   Dim ajob
   Dim strAbsFile
   Dim strFileExtension
   Dim objFSO
   Dim objFile
   Dim aauthorization
   Dim passwordcheck
   'Response.Write "<h1>" & CStr(Dession.aCommandTimeout) & "</h1>"
   'Response.End
   s=CStr(Request.ServerVariables("HTTP_COOKIE"))
   if len(s)=0 then
'$$<
'   startPage = "/icormanager/default.asp"
'   currentPage = Request.ServerVariables("SCRIPT_NAME")
'   if strcomp(currentPage,startPage,1) then
'      Response.Redirect(startPage)
'   end if
'$$>
      Response.Redirect "/icormanager/"
      Response.Flush
      Response.End
   end if
   Process = 1
'   OutputStr "ASP Process start: " & paction
'   OutputStr "  Dession(uid): " & CStr(Dession("uid"))
'   OutputStr "  Request.QueryString(relogin): " & CStr(Request.QueryString("relogin"))
'   OutputStr "  Request.Dession(relogin2): " & CStr(Dession("relogin2"))
'   OutputStr "  Request.ServerVariables(HTTP_AUTHORIZATION): " & aauthorization
   RefreshSession
   aauthorization=CStr(Request.ServerVariables("HTTP_AUTHORIZATION"))
   if aauthorization="" then
      Dession("uid")=-1
   end if
   'xx=Dession("uid")
   'response.write "XX: " & CStr(xx) & " type: " & CStr(typename(xx))
   'response.end
   duid=-1
   on error resume next
   duid=Dession("uid")
   if (TypeName(duid)<>"Integer") and (TypeName(duid)<>"Long") then
      duid=-1
      'Dession("uid")=-1
      'Response.Redirect "/icormanager/"
      'Response.End
         'response.write "<h1>" & CStr(TypeName(duid)) & "</h1>"
      'response.end
   end if
   'on error goto 0
   If (duid<0) or (Request.QueryString("relogin")="1") Then
      Dession("_passwordcheck")=0
      Dession("uid") = -1
      If ICORExists() < 0 Then
         Response.Write "<html><head><title>Zadanie dla administratora</title></head><body bgcolor=ivory>"
         Response.Write "<p align=""left""><font size=""5"" color=""#000080"">Serwer ICOR nie jest uruchomiony. Zalecany kontakt z administratorem.</font></p>"
         Response.Write "</body></html>"
         Exit Function
      End If
      errkind = 0 ' obsluga innych przegladarek
      If errkind < 0 Then
         Response.Write ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\DoProcess", Request.QueryString & SPLIT_CHAR_PARAM & GetServerVariables(), errkind, Request.Form, "")
         Exit Function
      End If
      if (aauthorization<>"") then
         Dession("uid") = ICORLogin()
      end if
      duid=Dession("uid")
      If duid > 0 Then
         Call GetICORSessionVars
         Dession("username")=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\GetUserByUID", "", -1, "", "")
         if Dession("UserImpersonated")=0 then
            ret=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Administration\User\PasswordCheckChange","",0,"","OK")
         else
            ret="OK"
         end if
         if ret="OK" then
            Dession("_passwordcheck")=99
            Response.Write ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\DoProcess", Request.QueryString & SPLIT_CHAR_PARAM & GetServerVariables(), 5, Request.Form, "")
         else
            Response.Write ret
            Dession("_passwordcheck")=1
         end if
      Else
         if Dession("relogin2")=1 then
            Dession("relogin2")=2
         elseif Dession("relogin2")<>2 then
            Dession("relogin2")=1
         end if
         Response.Buffer = True
         Response.Write "<html><head><title>Brak autoryzacji</title></head><body><h1><font color=red>Niepoprawna autoryzacja do serwera ICOR.</font></h1></body></html>"
         Response.Status = "401 Unauthorized"
         'OutputStr "Authorisation send!"
         Response.AddHeader "WWW-Authenticate","BASIC Realm=ICOR"
         Response.End
      End If
   Else
      'OutputStr "** Logged in! Dession(uid): " & CStr(Dession("uid"))
      Dession("relogin2")=0
      s=Dession("username")
      if (s="") or (s="Administrator") then
         Dession("username")=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\GetUserByUID", "", -1, "else 1", "")
      end if
      passwordcheck=Dession("_passwordcheck")
      if passwordcheck=0 then
         passwordcheck=99
      end if
      if passwordcheck=1 then
         ret=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Administration\User\PasswordCheckChange","",1,Request.Form,"OK")
         if ret="OK" then
            Dession("_passwordcheck")=99
            passwordcheck=99
            Response.Write ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\DoProcess", Request.QueryString & SPLIT_CHAR_PARAM & GetServerVariables(), 5, Request.Form, "")
         else
            Response.Write ret
         end if
      end if
      'response.write "1 " & Dession("username") & " " & Dession("_passwordcheck")
      'Response.End
      if passwordcheck=99 then
         ajob=Request.Querystring("jobtype")
         'OutputStr "ASP jobtype: " & ajob
         Select Case ajob
            Case "icorclose"
               Call ICORClose
               Response.Write "<html><body><h1>Shutdown successful.</h1><hr><button onclick='javascript:history.back();'> Return </button></body></html>"
               Response.End
            Case "lastvisithistory"
               aid=Request.Querystring("id")
               if aid<>"" then
                  Dession("lastvisithistory_ignore_click")=1
                  if aid="searchform" then
                     j=Dession("LastVisitHistoryID")-1
                     for i=1 to MAX_HISTORY
                        if j<1 then
                           j=MAX_HISTORY
                        end if
                        amode=Dession("lastvisithistory_mode_" & CStr(j))
                        aurl=Dession("lastvisithistory_url_" & CStr(j))
                        if aurl<>"" then
                           aid=CStr(j)
                        end if
                        if amode="searchform" then
                           exit for
                        end if
                        j=j-1
                     next
                  end if
                  if aid="singleobject" then
                     j=Dession("LastVisitHistoryID")-1
                     for i=1 to MAX_HISTORY
                        if j<1 then
                           j=MAX_HISTORY
                        end if
                        amode=Dession("lastvisithistory_mode_" & CStr(j))
                        aurl=Dession("lastvisithistory_url_" & CStr(j))
                        if aurl<>"" then
                           aid=CStr(j)
                        end if
                        if amode<>"searchform" then
                           exit for
                        end if
                        j=j-1
                     next
                  end if
                  aurl=Dession("lastvisithistory_url_" & aid)
                  Response.Write "<html xmlns=""http://www.w3.org/TR/REC-html40"">"
                  Response.Write "<head>"
                  Response.Write "<meta HTTP-EQUIV='Content-Type' content='text/html; charset=utf-8'>"
                  Response.Write "<meta http-equiv='Content-Language' content='pl'>"
                  Response.Write "<meta name='pragma' content='no-cache'>"
                  Response.Write "<meta name='keywords' content='ICOR object oriented database WWW information management repository'>"
                  Response.Write "<meta name='author' content='" & Application("META_AUTHOR") & "'>"
                  Response.Write "<META NAME='generator' CONTENT='ICOR'>"
                  Response.Write "<META HTTP-EQUIV='expires' CONTENT='Mon, 1 Jan 2001 01:01:01 GMT'>"
                  Response.Write "<base target='TEXT'>"
                  Response.Write "</head>"
                  Response.Write "<body>"
                  Response.Write "<form name=form1 id=form1 METHOD='post' ACTION='" & Dession("lastvisithistory_url_" & aid) & "'>"
                  Response.Write Dession("lastvisithistory_form_" & aid)
                  Response.Write "</form>"
                  Response.Write chr(10) & "<script language=javascript defer>" & chr(10)
                  Response.Write "form1.submit();" & chr(10)
                  Response.Write "</script>" & chr(10)
                  Response.Write "</body>"
                  Response.Write "</html>"
               else
                  Response.Write "<html xmlns=""http://www.w3.org/TR/REC-html40"">"
                  Response.Write "<head><link rel=STYLESHEET type='text/css' href='/icormanager/icor.css' title='SOI'>"
                  Response.Write "<meta HTTP-EQUIV='Content-Type' content='text/html; charset=utf-8'>"
                  Response.Write "<meta http-equiv='Content-Language' content='pl'>"
                  Response.Write "<meta name='pragma' content='no-cache'>"
                  Response.Write "<meta name='keywords' content='ICOR object oriented database WWW information management repository'>"
                  Response.Write "<meta name='author' content='" & Application("META_AUTHOR") & "'>"
                  Response.Write "<META NAME='generator' CONTENT='ICOR'>"
                  Response.Write "<META HTTP-EQUIV='expires' CONTENT='Mon, 1 Jan 2001 01:01:01 GMT'>"
                  Response.Write "<title>Ostatnio odwiedzane</title>"
                  Response.Write "<base target='TEXT'>"
                  Response.Write "</head>"
                  Response.Write "<body>"
                  Response.Write "<span class='objectsviewcaption'>Ostatnio odwiedzane:</span>"
                  Response.Write "<TABLE class='objectsviewtable'>"
                  Response.Write "<TR>"
                  Response.Write "   <TH class='objectsviewheader'>l.p.</TH>"
                  Response.Write "   <TH class='objectsviewheader'>Data i czas</TH>"
                  Response.Write "   <TH class='objectsviewheader'>Pozycja</TH>"
'                  Response.Write "   <TH class='objectsviewheader'>Warto��</TH>"
                  Response.Write "</TR>"
                  j=Dession("LastVisitHistoryID")-1
                  k=1
                  for i=1 to MAX_HISTORY
                     if j<1 then
                        j=MAX_HISTORY
                     end if
                     aurl=Dession("lastvisithistory_url_" & CStr(j))
                     if aurl<>"" then
                        aform=Dession("lastvisithistory_form_" & CStr(j))
'                        if aform<>"" then
                        aurl="icormain.asp?jobtype=lastvisithistory&id=" & CStr(j)
'                        end if
                        Response.Write "<TR class='objectsviewrow'>"
                        Response.Write "<td class='objectsviewdataeven' align=left valign=top nowrap><a class='objectitemasanchor' href='" & aurl & "'>" & CStr(k) & "</a></td>"
                        Response.Write "<td class='objectsviewdataeven' align=left valign=top nowrap><a class='objectitemasanchor' href='" & aurl & "'><font color=navy>" & Dession("lastvisithistory_time_" & CStr(j)) & "</font></a></td>"
'                        Response.Write "<td class='objectsviewdataeven' align=left nowrap><a class='objectitemasanchor' href='" & aurl & "'>" & Dession("lastvisithistory_name_" & CStr(j)) & "</a></td>"
'                        Response.Write "<td class='objectsviewdataeven' align=left nowrap><a class='objectitemasanchor' href='" & aurl & "'>" & Dession("lastvisithistory_value_" & CStr(j)) & "</a></td>"
                        Response.Write "<td class='objectsviewdataeven' align=left valign=top nowrap><a class='objectitemasanchor' href='" & aurl & "'><font color=navy>" & Dession("lastvisithistory_name_" & CStr(j)) & "</font></a>"
                        if Dession("lastvisithistory_value_" & CStr(j))<>"" then
                           Response.Write "<table>"
                           Response.Write Dession("lastvisithistory_value_" & CStr(j))
                           Response.Write "</table>"
                        end if
                        Response.Write "</td>"
                        Response.Write "</TR>"
                        k=k+1
                     end if
                     j=j-1
                  next
                  Response.Write "</TABLE>"
                  Response.Write "</body>"
                  Response.Write "</html>"
               end if
            Case Else
               'OutputStr "ASP paction: " & paction
               Response.Charset = "utf-8"
   '            Response.CodePage = 65001
               Select Case paction
                  Case "Introduction"
                     iaction = 1
                  Case "Contents"
                     iaction = 2
                  Case "NavigationBar"
                     iaction = 3
                  Case "NavigationButtons"
                     iaction = 4
                     Response.CacheControl = "Private"
                     Response.ExpiresAbsolute = #1/1/1999 1:50:00 AM#
                  Case "Default"
                     Call GetICORSessionVars
                     iaction = 5
                  'Case "PPCMain"
                     'iaction=7
                  Case Else
                     iaction = 0
                     Response.CacheControl = "no-cache"
                     Response.ExpiresAbsolute = #1/1/1999 1:50:00 AM#
                     Response.AddHeader "Pragma", "no-cache"
                     Response.Expires = -1
                     If GetInputValue("MIMEExcel") = "1" Then Response.ContentType = "application/vnd.ms-excel"
                     If GetInputValue("MIMEWord") = "1" Then Response.ContentType = "application/msword"
                     If (GetInputValue("MIMEXML") = "1") or (GetInputValue("XMLData") = "1") Then Response.ContentType = "text/xml"
                     If (GetInputValue("MIMEJSON") = "1") or (GetInputValue("XMLData") = "json") Then Response.ContentType = "application/json"
                     If GetInputValue("MIMEClass") <> "" Then Response.ContentType = GetInputValue("MIMEClass")
                     If GetInputValue("MIMEClass1") <> "" Then Response.ContentType = GetInputValue("MIMEClass1")
                     If GetInputValue("MIMEClass2") <> "" Then Response.ContentType = GetInputValue("MIMEClass2")
                     If GetInputValue("MIMEClass3") <> "" Then Response.ContentType = GetInputValue("MIMEClass3")
                     If GetInputValue("MIMEClass4") <> "" Then Response.ContentType = GetInputValue("MIMEClass4")
                     If GetInputValue("MIMEClass5") <> "" Then Response.ContentType = GetInputValue("MIMEClass5")
'                     If GetInputValue("MIMESave") = "1" Then Response.ContentType = "application/save"
                     If GetInputValue("MIMESave") = "1" Then Response.ContentType = "application/force-download"
'                     If ((LCase(Response.ContentType) = "application/save") or (LCase(Response.ContentType) = "application/force-download")) And (GetInputValue("MIMEContentFileName") <> "") Then
                     If GetInputValue("MIMEContentFileName") <> "" Then
                        Response.AddHeader "content-disposition", "attachment; filename=""" & GetInputValue("MIMEContentFileName") & """"
                     End if
               End Select
               s=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\DoProcess", Request.QueryString & SPLIT_CHAR_PARAM & GetServerVariables(), iaction, Request.Form, "")
               alen=len(s)
               i=1
               do while i<=alen
                  Response.Write Mid(s,i,30000)
                  Response.Flush
                  i=i+30000
               loop
         End Select
      End If
   End If
End Function
%>