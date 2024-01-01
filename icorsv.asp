<%@  codepage="65001" language="VBSCRIPT" %>
<!-- #include file="all.asp" -->
<!--#include file ="definc_createIISI.asp"-->
<%

NoCache

amode=ReplaceIllegalCharsLev0Text(mid(Request.QueryString("m"),1,80))
if amode="" then
   amode=ReplaceIllegalCharsLev0Text(mid(Request.Form("m"),1,80))
end if

'#####################################################################################
if amode="evServerStatus" then
   Response.ContentType = "application/json"
   on error resume next
   ret=GetICORServerStatus()
   If (Err.Number<>0) or (ret="") Then
      s=Err.Description
      response.write "{""status"":1,""info"":""ICOR EServer is not responding"",""errinfo"":""" & Replace(s,"""","'") & """}"
      Err.Clear
      Response.End
   End If
   on error goto 0
   Response.write ret
   Response.End
end if

'#####################################################################################

if amode="evDashboardStart" then
   Response.Write "<dane><info>ok</info></dane>"
   Response.end
end if

if amode="evICORStart" then
   Response.Write "<dane><info>ok</info></dane>"
   Response.end
end if

if amode="evUserDataLastLogins" then
   dim ret
   ret=Proc_ExecuteMethodRetS("py25","CLASSES\Library\NetBase\WWW\Administration\User\WWWGetUserLog", "", 10, "JSON", "")
   Response.ContentType = "application/json"
   response.Write ret
end if

'#####################################################################################

if amode="evLetterSearch" then
   aletter=ReplaceIllegalCharsLev0Text(URL.decode(mid(Request.QueryString("l"),1,2)))
   if aletter="" then
      Response.End
   end if
   asql=""
   asql=asql & "select ( "
   asql=asql & "select "
   asql=asql & "	_OID "
   asql=asql & "      ,Haslo "
   asql=asql & "      ,Rodzaj "
   asql=asql & "  FROM V01_BZR_86094 "
   asql=asql & "where Jezyk='pl' "
   asql=asql & "and Status in ('A','I') "
   if aletter="_" then
      asql=asql & "and Haslo like '[^A-Z]%' "
   else
      asql=asql & "and Haslo like '" & aletter & "%' "
   end if
   asql=asql & "order by Haslo "
   asql=asql & "for xml path('haslo'), ROOT('hasla'), ELEMENTS "
   asql=asql & ") as result "
   call ExecuteSQLXML(asql)
end if

if amode="evTextSearch" then
   t=ReplaceIllegalCharsLev0Text(URL.decode(mid(Request.QueryString("t"),1,40)))
   f=ReplaceIllegalCharsLev0Text(URL.decode(mid(Request.QueryString("f"),1,10)))
   if t="" then
      Response.End
   end if
   asql=""
   asql=asql & "select ( "
   asql=asql & "select top 200 "
   asql=asql & "	_OID "
   asql=asql & "      ,Haslo "
   asql=asql & "      ,Rodzaj "
   asql=asql & "  FROM V01_BZR_86094 "
   asql=asql & "where Jezyk='pl' "
   asql=asql & "and Status in ('A','I') "
   if f="1" then
      asql=asql & "and Tresc like '%" & t & "%' "
   else 
      asql=asql & "and Haslo like '%" & t & "%' "
   end if
   asql=asql & "order by Haslo "
   asql=asql & "for xml path('haslo'), ROOT('hasla'), ELEMENTS "
   asql=asql & ") as result "
   call ExecuteSQLXML(asql)
end if

if amode="evItemClick" then
   ioid=GetSafeOID(Request.QueryString("i"))
   if ioid="" then
      Response.End
   end if 
   asql=""
   asql=asql & "select ( "
   asql=asql & "select  "
   asql=asql & "	th._OID "
   asql=asql & "    ,th.Haslo "
   asql=asql & "    ,th.Rodzaj "
   asql=asql & "    ,th.Tresc "
   asql=asql & "    ,th.Jezyk, "
   asql=asql & "	(select  "
   asql=asql & "		th1._OID "
   asql=asql & "		,th1.Haslo "
   asql=asql & "		,th1.Rodzaj "
   asql=asql & "		,th1.Jezyk "
   asql=asql & "	from V01_BZR_86095 tp1  "
   asql=asql & "	left join V01_BZR_86094 th1 on tp1.Haslo2=th1._OID "
   asql=asql & "	where  "
   asql=asql & "		tp1.Haslo1=th._OID "
   asql=asql & "		and tp1.Rodzaj like 'T_umaczenie' "
   asql=asql & "	order by th1.Jezyk,th1.Haslo asc "
   asql=asql & "	for xml path('tlumaczenie'),type) as 'tlumaczenia', "
   asql=asql & "	(select  "
   asql=asql & "		th1._OID "
   asql=asql & "		,th1.Haslo "
   asql=asql & "		,th1.Rodzaj "
   asql=asql & "		,th1.Jezyk "
   asql=asql & "	from V01_BZR_86095 tp1  "
   asql=asql & "	left join V01_BZR_86094 th1 on tp1.Haslo2=th1._OID "
   asql=asql & "	where  "
   asql=asql & "		tp1.Haslo1=th._OID "
   asql=asql & "		and tp1.Rodzaj='Referencja' "
   asql=asql & "	order by th1.Haslo asc "
   asql=asql & "	for xml path('referencja'),type) as 'referencje', "
   asql=asql & "	(select  "
   asql=asql & "		th1._OID "
   asql=asql & "		,th1.Haslo "
   asql=asql & "		,th1.Rodzaj "
   asql=asql & "		,th1.Jezyk "
   asql=asql & "	from V01_BZR_86095 tp1  "
   asql=asql & "	left join V01_BZR_86094 th1 on tp1.Haslo2=th1._OID "
   asql=asql & "	where  "
   asql=asql & "		tp1.Haslo1=th._OID "
   asql=asql & "		and tp1.Rodzaj='Synonim' "
   asql=asql & "	order by th1.Haslo asc "
   asql=asql & "	for xml path('synonim'),type) as 'synonimy', "
   asql=asql & "	(select  "
   asql=asql & "		th1._OID "
   asql=asql & "		,th1.Haslo "
   asql=asql & "		,th1.Rodzaj "
   asql=asql & "		,th1.Jezyk "
   asql=asql & "	from V01_BZR_86095 tp1  "
   asql=asql & "	left join V01_BZR_86094 th1 on tp1.Haslo2=th1._OID "
   asql=asql & "	where  "
   asql=asql & "		tp1.Haslo1=th._OID "
   asql=asql & "		and tp1.Rodzaj like 'Skr_t' "
   asql=asql & "	order by th1.Haslo asc "
   asql=asql & "	for xml path('skrot'),type) as 'skroty', "
   asql=asql & "	(select  " 
   asql=asql & "		th1._OID "
   asql=asql & "		,th1.Haslo "
   asql=asql & "		,th1.Rodzaj "
   asql=asql & "		,th1.Jezyk "
   asql=asql & "	from V01_BZR_86095 tp1  "
   asql=asql & "	left join V01_BZR_86094 th1 on tp1.Haslo2=th1._OID "
   asql=asql & "	where  "
   asql=asql & "		tp1.Haslo1=th._OID "
   asql=asql & "		and tp1.Rodzaj='Akronim' "
   asql=asql & "	order by th1.Haslo asc "
   asql=asql & "	for xml path('akronim'),type) as 'akronimy', "
   
   asql=asql & "( "
   asql=asql & "	select top 15 "
   asql=asql & "		ti.Haslo, "
   asql=asql & "		ti.Obiekt, "
   asql=asql & "		tf.ChapterID, "
   asql=asql & "		(select sum(Licznik) from v01_bzr_86093 tl where tl.obiekt=ti.Obiekt) Licznik, "
   asql=asql & "		tt.Tytul "
   asql=asql & "	from v01_bzr_86097 ti "
   asql=asql & "	left join V01_FTSSEARCH_0 tf on ti.Obiekt=tf.OIDRef "
   asql=asql & "	left join ( "
   asql=asql & "		select _OID,Tytul from v01_bzr_86032 ta1  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86044 ta2  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86031 ta3  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86035 ta4  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86037 ta5  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86036 ta6  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86038 ta7  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86033 ta8  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86034 ta9  "
   asql=asql & "		union all "
   asql=asql & "		select _OID,Tytul from v01_bzr_86043 ta10 "
   asql=asql & "	) tt on tt._OID=ti.Obiekt "
   asql=asql & "	where  "
   asql=asql & "		tf.TableID in (86032,86044,86031,86035,86037,86036,86038,86033,86034,86043) and "
   asql=asql & "		ti.haslo=th._OID "
   asql=asql & "	order by Licznik desc "
   asql=asql & "	for xml path('artykul'),type) as 'artykuly' "
   
   asql=asql & "  FROM V01_BZR_86094 th "
   asql=asql & "where "
   asql=asql & "th.Status in ('A','I') "
   asql=asql & "and th._OID='" & ioid & "' "
   asql=asql & "for xml path('haslo'), ROOT('hasla'), ELEMENTS "
   asql=asql & ") as result "
   call ExecuteSQLXML(asql)
end if

if amode="evChapterView" then
   ioid=GetSafeOID(CStr(Request.QueryString("ioid")))
   achapter=ReplaceIllegalCharsLev0Text(mid(Request.QueryString("chapter"),1,12))
   if achapter="" then
      Response.End
   end if

   astatus=""
   acnt=0
   
   call db.Open()
   asql="select * from V01_BZR_86093 WHERE Rodzaj='CC' and Rozdzial='" & achapter & "' and Obiekt='" & ioid & "'"
   set rs=db.GetRS(asql)
   if rs.EOF or rs.BOF then
      rs.AddNew
      rs("Rodzaj")="CC"
      rs("Rozdzial")=achapter
      rs("Obiekt")=ioid
      rs("Licznik")=1
      acnt=1
   else 
      acnt=1+rs("Licznik")
      rs("Licznik")=acnt
   end if
   rs.Update
   astatus="OK"
   call db.CloseRS(rs)
   call db.Close()
   
   response.write "{""status"":""" & astatus & """,""cnt"":""" & CStr(acnt) & """}"
end if

if amode="taField" then
   f=ReplaceIllegalCharsLev0Text(mid(Request.Form("f"),1,40))
   afield=ReplaceIllegalCharsLev0Text(mid(Request.Form("field"),1,200))
   aquery=ReplaceIllegalCharsLev0Text(mid(Request.Form("query"),1,40))
   avalue=ReplaceIllegalCharsLev0Text(mid(Request.Form("value"),1,40))
   Response.ContentType="application/json"

   litems=""
   if f="dict" then
      if afield="search01" then
         call db.Open()
         asql="select distinct top 15 Haslo from v01_bzr_86094 where Jezyk='pl' and Status in ('A','I') and Haslo like '%" & aquery & "%' order by Haslo"
         set rs=db.GetRS(asql)
         do while not rs.EOF and not rs.BOF
            if litems<>"" then
               litems=litems & ","
            end if
            avalue=GetRSValue(rs,"Haslo")
            litems=litems+""""+JSONEncode(avalue)+""""
            rs.MoveNext
            Loop
         call db.CloseRS(rs)
      end if
   end if
   
   Response.Write "[" & litems & "]"
end if
%> 