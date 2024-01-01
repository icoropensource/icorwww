<%
Function GetSQLCharStringAsString(data)
   dim myRegExp,matches,myMatch,i,avalue,ivalue
   Set myRegExp = New RegExp
   myRegExp.IgnoreCase = True
   myRegExp.Global = True
   data=replace(data,chr(10)," ")
   data=replace(data,chr(13)," ")
   myRegExp.Pattern = "[ \+]*\([ \+]*"
   data=myRegExp.Replace(data,"(")
   myRegExp.Pattern = "[ \+]*\)[ \+]*"
   data=myRegExp.Replace(data,")")
   myRegExp.Pattern = "[ \+]*char\("
   data=myRegExp.Replace(data,"char(")
   myRegExp.Pattern = "char\(([0-9]+)\)"
   set matches=myRegExp.Execute(data)
   For Each myMatch in matches
      For i=0 to myMatch.SubMatches.Count-1
         avalue=myMatch.SubMatches(i)
         ivalue=CLng(avalue)
         if ivalue<128 then
            data=replace(data,"char("+avalue+")",chr(ivalue))
         end if
      Next
   Next
   GetSQLCharStringAsString=data
End Function

Function GetSecurityScore(aline)
   dim l_blacklist_score_51,l_blacklist_score_21,l_blacklist_score_11,iCounter
   l_blacklist_score_51=Array("--","@@"," char","+char","varchar","declare","coalesce","sysobjects", "syscolumns",_
"sql_server","error_converting","type_mismatch","incorrect_syntax","unclosed_quotation","conversion_failed",_
"information_schema",".columns","data_type","column_name",".tables","table_name","+or+1","+and+1","/*", "*/")
   l_blacklist_score_21=Array("select","cast","delete","drop","table","'","%27","%10","insert","update","union","%u","\u","&#","u&","\x","master")
   l_blacklist_score_11=Array(";", "@","nchar","alter", "begin", "create", "cursor","exec","execute", "fetch", "kill", "open")
   GetSecurityScore=0
   if len(aline)<2 then
      Exit Function
   end if
   aline=LCase(aline)
   aline=GetSQLCharStringAsString(aline)
   For iCounter=0 to uBound(l_blacklist_score_51)
      if InStr(aline,l_blacklist_score_51(iCounter))>0 then
         GetSecurityScore=GetSecurityScore+51
      end if
   Next
   For iCounter=0 to uBound(l_blacklist_score_21)
      if InStr(aline,l_blacklist_score_21(iCounter))>0 then
         GetSecurityScore=GetSecurityScore+21
      end if
   Next
   For iCounter=0 to uBound(l_blacklist_score_11)
      if InStr(aline,l_blacklist_score_11(iCounter))>0 then
         GetSecurityScore=GetSecurityScore+11
      end if
   Next
End Function

Function CheckSecurityScore(aline)
   dim l_blacklist_score_51,l_blacklist_score_21,l_blacklist_score_11,iCounter,ascore
   l_blacklist_score_51=Array("--","@@"," char","+char","varchar","declare","coalesce","sysobjects", "syscolumns",_
"sql_server","error_converting","type_mismatch","incorrect_syntax","unclosed_quotation","conversion_failed",_
"information_schema",".columns","data_type","column_name",".tables","table_name","+or+1","+and+1","/*", "*/")
   l_blacklist_score_21=Array("select","cast","delete","drop","table","'","%27","%10","insert","update","union","%u","\u","&#","u&","\x","master")
   l_blacklist_score_11=Array(";", "@","nchar","alter", "begin", "create", "cursor","exec","execute", "fetch", "kill", "open")
   if len(aline)<2 then
      CheckSecurityScore=True
      Exit Function
   end if
   CheckSecurityScore=False
   ascore=0
   aline=LCase(aline)
   aline=GetSQLCharStringAsString(aline)
   For iCounter=0 to uBound(l_blacklist_score_51)
      if InStr(aline,l_blacklist_score_51(iCounter))>0 then
         Exit Function
      end if
   Next
   For iCounter=0 to uBound(l_blacklist_score_21)
      if InStr(aline,l_blacklist_score_21(iCounter))>0 then
         ascore=ascore+21
         if ascore>50 then
            Exit Function
         end if
      end if
   Next
   For iCounter=0 to uBound(l_blacklist_score_11)
      if InStr(aline,l_blacklist_score_11(iCounter))>0 then
         ascore=ascore+11
         if ascore>50 then
            Exit Function
         end if
      end if
   Next
   CheckSecurityScore=True
End Function

Function IsValidHex(aval)
   dim c1,i,w1
   IsValidHex=True
   for i=1 to len(aval)
      c1=mid(aval,i,1)
      w1=(c1>="a" and c1<="f") or (c1>="0" and c1<="9")
      if not w1 then
         IsValidHex=False
         Exit Function
      end if
   next
End Function

Function URLDecode(What)
  Dim Pos, pPos, s1
  On Error Resume Next
  Dim Stream
  Set Stream = Server.CreateObject("ADODB.Stream")
  If 0 = Err.Number Then ' decode URL data using ADODB.Stream, If possible
    On Error GoTo 0
    Stream.Type = 2 'String
    Stream.Open

    Pos = InStr(1, What, "%") ' replace all %XX to character
    pPos = 1
    Do While Pos > 0
      Stream.WriteText Mid(What, pPos, Pos - pPos)
      s1=Mid(What, Pos + 1, 2)
      if IsValidHex(s1) then
         Stream.WriteText Chr(CLng("&H" & s1))
      else
         Stream.WriteText "%" & s1
      end if
      pPos = Pos + 3
      Pos = InStr(pPos, What, "%")
    Loop
    Stream.WriteText Mid(What, pPos)

    Stream.Position = 0 ' read the text stream
    URLDecode = Stream.ReadText
    Stream.Close  ' free resources
  Else ' URL decode using string concentation
    On Error GoTo 0
    ' Slow, do not use it for data length over 100k
    Pos = InStr(1, What, "%")
    Do While Pos>0 
      What = Left(What, Pos-1) + _
        Chr(Clng("&H" & Mid(What, Pos+1, 2))) + _
        Mid(What, Pos+3)
      Pos = InStr(Pos+1, What, "%")
    Loop
    URLDecode = What
  End If
End Function

Function URLDecodePlus(What)
  What=Replace(What, "+", " ")
  URLDecodePlus=URLDecode(What)
End Function

Function URLDecodeFull(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = eval("&h" + Mid(enStr, i + 1, 2))
            If v < 128 Then
                deStr=deStr&chr(v)
                i = i + 2
            Else
                If isvalidhex(Mid(enStr, i, 3)) Then
                    If isvalidhex(Mid(enStr, i + 3, 3)) Then
                        v = eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr=deStr&chr(v)
                        i = i + 5
                    Else
                        v = eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr=deStr&chr(v)
                        i = i + 3
                    End If
                Else
                    destr=destr&c
                End If
            End If
        Else
            deStr=deStr&c
        End If
    Next
    URLDecodeFull = deStr
End Function

Function URLDecodeFullPlus(What)
  What=Replace(What, "+", " ")
  URLDecodeFullPlus=URLDecodeFull(What)
End Function

Function GetSecurityScoreURL(aline)
   aline=URLDecode(aline)
   GetSecurityScoreURL=GetSecurityScore(aline)
End Function

Function CheckSecurityScoreURL(aline)
   aline=URLDecode(aline)
   CheckSecurityScoreURL=CheckSecurityScore(aline)
End Function

'arequest=CStr(Request.QueryString)
'Response.Write "<p> Request: " & arequest & "</p>"
'Response.Write "<p> Score: " & GetSecurityScoreURL(arequest) & "</p>"
'Response.Write "<p> Check: " & CStr(CheckSecurityScoreURL(arequest)) & "</p>"
%>