<%
dim Dession

Class DBSession
   Public aCommandTimeout,aConnectionTimeout,sessionToken
   Public cn,rs
   Private Sub Class_Initialize()
      me.aCommandTimeout=15
      me.aConnectionTimeout=15
      me.sessionToken=""
   End Sub
   Private Sub Class_Terminate()
      Me.Close()
   End Sub

   Public Default Property Get Item(ByVal Key)
      dim stoken,tname,svalue,rs,vempty,asql
      Key=LCase(Key)
      if Key="ui_skin" then
         Item="redmond"
         exit property
      end if
      stoken=me.GetSessionToken()
      if stoken="" then
         Item=""
         exit property
      end if
      Item=vempty
      asql="select * from icor.sessionvalues where tokenid='" & stoken & "' and valuename='" & Replace(Key,"'","''") & "';"
      set rs=me.GetRs(asql)
      if not rs.EOF and not rs.BOF then
         tname=CStr(rs("valuetype"))
         Item=CStr(rs("value"))
         if tname="Byte" then
            Item=CByte(Item)
         end if
         if tname="Integer" then
            Item=CLng(Item)
         end if
         if tname="Long" then
            Item=CLng(Item)
         end if
         if tname="Single" then
            Item=CSng(Item)
         end if
         if tname="Double" then
            Item=CDbl(Item)
         end if
         if tname="Currency" then
            Item=CCur(Item)
         end if
         if tname="Decimal" then
            Item=CDbl(Item) '???
         end if
         if tname="Date" then
            Item=CDate(Item)
         end if
         if tname="String" then
            Item=CStr(Item)
         end if
         if tname="Boolean" then
            Item=CBool(Item)
         end if
         if tname="Empty" then
            Item=vempty
         end if
      end if
      'response.write "<h1>dessionGET: " & CStr(key) & " : " & CStr(len(Item)) & " - " & Server.HTMLEncode(Item) & " $$</h1>"
      call me.CloseRs(rs)
   End Property
   
   Public Property Let Item(ByVal Key, ByVal Value)
      dim stoken
      dim tname,w,rs
      dim rcnt
      Key=LCase(Key)
      stoken=me.GetSessionToken()
      if stoken="" then
         exit property
      end if
      tname=TypeName(Value)
      w=False
      if tname="Byte" or tname="Integer" or tname="Long" then
         w=True
      end if
      if tname="Single" or tname="Double" then
         w=True
      end if
      if tname="Currency" or tname="Decimal" then
         w=True
      end if
      if tname="Date" then
         w=True
      end if
      if tname="String" then
         w=True
      end if
      if tname="Boolean" then
         w=True
      end if
      if tname="Empty" then
         w=True
         Value=""
      end if
      if tname="IStringList" then
         tname="String"
         w=True
      end if
      'Null <object type> Object Unknown Nothing Error
      if not w then
         Response.Write "BAD_TYPE: " & Key & " : " & tname
         Response.End
         Exit Property
      end if
      asql="INSERT INTO icor.sessionvalues (tokenid, status, valuename, valuetype, value, _datetime) VALUES ( '" & CStr(stoken) & "' , 'A' , $DATA1$" & CStr(Key) & "$DATA1$ , $DATA2$" & CStr(tname) & "$DATA2$ , $DATAX$" & CStr(Value) & "$DATAX$ , current_timestamp )     ON CONFLICT (tokenid,valuename) DO UPDATE SET valuetype=EXCLUDED.valuetype, value=EXCLUDED.value, _datetime=EXCLUDED._datetime; "
      on error resume next
      Me.cn.Execute asql
      If Err Then
         Response.Write "<h1>Err1: " & Server.URLEncode(Err.Description) & "</h1>"
         Err.Clear
         response.write "<h1>len: " & len(Server.URLEncode(Value)) & " type: " & tname & " $$</h1>"
         If Err Then
            Response.Write "<h1>Err2: " & Server.URLEncode(Err.Description) & "</h1>"
            Err.Clear
         end if
         'LogLine(Value)
         response.write "<h1>value: " & CStr(Server.URLEncode(Value)) & " $$</h1>"
         response.write "<h1>value: " & CStr(Server.URLEncode(Value)) & " $$</h1>"
         If Err Then
            Response.Write "<h1>Err3: " & Server.URLEncode(Err.Description) & "</h1>"
            Err.Clear
         end if
         response.write "<h1>sql: " & Server.HTMLEncode(asql) & " $$</h1>"
         If Err Then
            Response.Write "<h1>Err4: " & Server.URLEncode(Err.Description) & "</h1>"
            Err.Clear
         end if
         response.write "<h1>.......</h1>"
         Response.End()
      End If
      On Error GoTo 0
   End Property

   Public Property Get SessionID()
      SessionID=me.GetSessionToken()
   End Property
   
   public function CheckValidSession(stoken)
      dim asql,rs
      CheckValidSession=False
      stoken=me.GetSafeOID(stoken)
      if stoken="" then
         exit function
      end if
      asql="select * from icor.sessiontokens where _oid='" & stoken & "' and status='A';"
      set rs=me.GetRs(asql)
      if not rs.EOF and not rs.BOF then
         CheckValidSession=True
      end if
      call me.CloseRs(rs)
   End Function

   public function GetSessionToken()
      dim sv,asql,rs
      GetSessionToken=me.sessionToken
      if GetSessionToken<>"" then
         exit function
      end if
      'sv=CStr(Request.Cookies("_isid"))
      'Response.Write "<h1>Cookie_isid=[" & sv & "]</h1>"
      dim scookies
      scookies=CStr(Request.ServerVariables("HTTP_COOKIE"))
      'Response.Write "<h1>Header=[" & scookies & "]</h1>"
      Dim i,aitems,avalues
      dim sitem,scookiename,scookievalue
      aitems=Split(scookies,";")
      for i=ubound(aitems) to lbound(aitems) step -1
         sitem=trim(aitems(i))
         avalues=Split(sitem,"=")
         scookiename=avalues(lbound(avalues))
         scookievalue=avalues(ubound(avalues))
         if (scookiename="_isid") and (len(scookievalue)=32) then
            'Response.Write "<h1>singlecookie=[" & scookiename & "=" & scookievalue & "]</h1>"
            if me.CheckValidSession(scookievalue) then
               'Response.Write "<h1>proper session=[" & scookievalue & "]</h1>"
               me.sessionToken=scookievalue
               GetSessionToken=scookievalue
               exit function
            else
               ' Response.AddHeader "Set-Cookie","_isid=; HttpOnly;"
               'Response.Write "<h1>empty session=[" & scookievalue & "]</h1>"
               Response.AddHeader "Set-Cookie","_isid=;"
            end if
         end if
      next
      asql="select * from icor.sessiontokens where _oid='-1';"
      set rs=me.GetRs(asql)
      rs.AddNew
      rs("status")="A"
      rs.Update
      GetSessionToken=CStr(rs("_OID"))
      call me.CloseRs(rs)
      me.sessionToken=GetSessionToken
      'Response.AddHeader "Set-Cookie","_isid=" & GetSessionToken & "; HttpOnly;"
      Response.AddHeader "Set-Cookie","_isid=" & GetSessionToken & ";"
   end function

   Function GetSafeOID(sInput)
      Dim iCounter,achar,wcheck
      GetSafeOID=""
      if len(sInput)<>32 then
         exit function
      end if
      for iCounter=1 to 32
         achar=mid(sInput,iCounter,1)
         if not ((achar>="0" and achar<="9") or (achar>="a" and achar<="f") or (achar>="A" and achar<="F")) then
            exit function
         end if
      next
      GetSafeOID=sInput
   End function
   Sub Open()
      '$$On Error Resume Next
      Set Me.cn = Server.CreateObject("ADODB.Connection")
      Me.cn.CursorLocation=3 'adUseClient
      Me.cn.CommandTimeout=me.aCommandTimeout
      Me.cn.ConnectionTimeout=me.aConnectionTimeout
      Me.cn.Open Application("CONNECTION_STRING_PG")
      If Err Then
         Set Me.cn = Nothing
         Err.Clear
         Response.End()
      End If
      '$$On Error GoTo 0
   End Sub
   Sub Close()
      '$$On Error Resume Next
      Me.cn.Close
      Set Me.cn = Nothing
      Err.Clear
      '$$On Error Goto 0
   End Sub
   Sub CloseRs(ars)
      If IsObject(ars) Then
         '$$On Error Resume Next
         if ars.State<>0 then
            ars.Close
         end if
         Set ars = Nothing
         '$$On Error Goto 0
      End If
   End Sub
   Function GetRs(asql)
      dim axrs
      Set axrs = Server.CreateObject("ADODB.Recordset")
      axrs.ActiveConnection = Me.cn
'Const adOpenForwardOnly = 0
'Const adOpenKeyset = 1
'Const adOpenDynamic = 2
'Const adOpenStatic = 3
      axrs.CursorType = 1
'adLockReadOnly = 1
'adLockPessimistic = 2
'adLockOptimistic = 3
'adLockBatchOptimistic = 4
      axrs.LockType = 3
      axrs.Source = asql
      axrs.Open
      Set GetRs=axrs
   End Function
End Class 

set Dession=new DBSession
call Dession.Open()
%>