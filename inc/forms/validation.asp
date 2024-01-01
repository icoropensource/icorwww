<script language="javascript" runat="server">
URL = {
   encode : function(s){return encodeURIComponent(s).replace(/'/g,"%27").replace(/"/g,"%22")},
   decode : function(s){return decodeURIComponent(s.replace(/\+/g," "))},
   encodeSimple : function(s){return s.replace(/ /g,"%20").replace(/\&/g,"%26").replace(/\=/g,"%3D")}
};
</script><%
Public function IsEmail(strEmail)
    Dim strUser, strHost, strArray, strExtensions, iCounter, iCounter2, blnFlag
    IsEmail = False
    blnFlag = False
    'Checking for illegal characters
    For iCounter = 1 To Len(strEmail)
    Select Case Asc(Mid(strEmail, iCounter, 1))
    'Checking Alpha, Numeric, "-", ".", "_"
    Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90, _
               97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122, _
               45,46,48,49,50,51,52,53,54,55,56,57,95
    Case 64
    'If "@" Then making sure it is the only 
    '     one
    if blnFlag = True Then
    Exit function
    Else: blnFlag = True
    End if
    Case Else
    Exit function
    End Select
    Next
    'Parsing the email into 2 parts
    if InStr(strEmail, "@") Then
    strUser = Left(strEmail, InStr(strEmail, "@") - 1)
    strHost = Mid(strEmail, InStr(strEmail, "@") + 1)
    Else: Exit function
    End if
    'Checking to make sure that there is at 
    '     least one "." in the host
    if Not InStr(strHost, ".") > 1 Then Exit function
      
      'Making sure that there is Not an "_" underscore in the domain
      if InStr(strHost, "_") > 1 Then Exit function
    'Making sure the first character is not 
    '     an "_" underscore
    if Left(strUser, 1) = "_" Then Exit function
    strArray = Split(strUser & "." & strHost, ".")
    For iCounter = 0 To UBound(strArray)
    'Making sure there are no ".." entered
    if Trim(strArray(iCounter)) = "" Then
    Exit function
    End if
    'Checking the domain extension
    if iCounter = UBound(strArray) Then
    'Creating array of legal domain extensio
    '     ns to validate with

'    strExtensions = Split("com;edu;gov;int;net;org;mil;info;cc;biz;name;pro;at;au;be;bg;br;by;ca;ch;cn;cy;cz;de;dk;ee;es;fi;fr;fx;gb;gr;hr;hu;ie;it;li;lt;lv;md;mt;mx;nl;no;nz;pe;pl;pt;ro;ru;se;si;sk;tr;ua;uk;us", ";")
    strExtensions = Split("com|net|org|edu|int|biz|de|fr|hu|it|nl|pl|ru|se|ua|uk|us|yy", "|") 'kompatybilne z /icorlib/validation/email_check.js

    'Below is a list of more country extensi
    '     ons i just find they bring in too many f
    '     alse positives.
    'Just paste them in the Split() line abo
    '     ve if you want to use them.
    ';af;al;dz;as;ad;ao;ai;aq;ag;ar;am;aw;ac
    '     ;at;au;az;bs;bh;bd;bb;by;be;bz;bj;bm;bt;
    '     bo;ba;bw;bv;br;io;bn;bg;bf;bi;kh;cm;ca;c
    '     v;ky;cf;td;gg;je;cl;cn;cx;cc;co;km;cg;ck
    '     ;cr;ci;hr;cu;cy;cz;dk;dj;dm;do;tp;ec;eg;
    '     sv;gq;er;ee;et;fk;fo;fj;fi;fr;fx;pf;tf;g
    '     a;gf;gm;ge;de;gh;gi;gr;gl;gd;gp;gu;gt;gn
    '     ;gw;gy;ht;hm;hn;hk;hu;is;in;id;ir;iq;ie;
    '     im;il;it;jm;jp;jo;kz;ke;ki;kp;kr;kw;kg;l
    '     a;lv;lb;ls;lr;ly;li;lt;lu;mo;mk;mg;mw;my
    '     ;ml;mt;mh;mq;mr;mu;mv;yt;mx;fm;md;mc;mn;
    '     ms;ma;mz;mm;na;nr;np;nl;an;nc;nz;ni;ne;n
    '     g;nu;nf;mp;no;om;pk;pw;pa;pg;py;pe;ph;pn
    '     ;pl;pt;pr;qa;re;ro;ru;rw;kn;lc;vc;ws;sm;
    '     st;sa;sn;sc;sl;sg;sk;si;sb;so;za;gs;es;l
    '     k;sh;pm;sd;sr;sj;sz;se;ch;sy;tw;tj;tz;th
    '     ;tg;tk;to;tt;tn;tr;tm;tc;tv;ug;ua;ae;uk;
    '     us;um;uy;uz;vu;va;ve;vn;vg;vi;wf;eh;ye;y
    '     u;zr;zm;zw
    For iCounter2 = 0 To UBound(strExtensions)
    if strArray(iCounter) = strExtensions(iCounter2) Then
    IsEmail = True
    Exit function
    End if
    Next
    End if
    Next
End function

Function JSONEncode(s)
   JSONEncode=Server.HTMLEncode(s)
   JSONEncode=replace(JSONEncode,"\","\\")
   JSONEncode=replace(JSONEncode,"/","\/")
   JSONEncode=replace(JSONEncode,chr(9),"&#9;")
   JSONEncode=replace(JSONEncode,chr(13),"")
   JSONEncode=replace(JSONEncode,chr(10),"&#10;")
End function


function UTF2Win(s)
   UTF2Win=s
end function

Function ReplaceIllegalCharsLev0(sInput)
   Dim sBadChars, iCounter
   sInput=UTF2Win(replace(sInput,"\","/"))
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "#", "%", "&", "'", "(", ")", ":", ";", "<", ">", "=", "[", "]", "?", "`", "|" )
   For iCounter=0 to uBound(sBadChars)
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next
   ReplaceIllegalCharsLev0=sInput
End function

Function ReplaceIllegalCharsLev0Text(sInput)
   Dim sBadChars, iCounter
   sInput=UTF2Win(replace(sInput,"\","/"))
   sInput=replace(sInput,"'","""")
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "#", "%", "<", ">" )
   For iCounter=0 to uBound(sBadChars)
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next
   ReplaceIllegalCharsLev0Text=sInput
End function

Function ReplaceIllegalCharsLev1JSON(sInput)
   Dim sBadChars, iCounter
   sInput=UTF2Win(replace(sInput,"\","/"))
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "#", "%", "&", "'", ";", "<", ">", "=", "[", "]", "`", "|", chr(13))
   For iCounter=0 to uBound(sBadChars)
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next
   sInput=replace(sInput,chr(10),"\n")
   ReplaceIllegalCharsLev1JSON=sInput
End function

Function ReplaceIllegalCharsLev1Markdown(sInput)
   Dim sBadChars, iCounter
   sInput=UTF2Win(replace(sInput,"\u","/u"))
   sBadChars=array("select", "drop", "insert", "delete", "xp_", "truncate", "update")
   For iCounter=0 to uBound(sBadChars)
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next
   sInput=replace(sInput,"&","&amp;")
   sInput=replace(sInput,"<","&lt;")
   sInput=replace(sInput,"""","&quot;")
   sInput=replace(sInput,">","&gt;")
   ReplaceIllegalCharsLev1Markdown=sInput
End function

Function ReplaceIllegalCharsLev0Password(sInput)
   Dim sBadChars, iCounter
   sInput=UTF2Win(sInput)
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", "'", """", "<", ">" )
   For iCounter=0 to uBound(sBadChars)
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next
   ReplaceIllegalCharsLev0Password=sInput
End function

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

Function stripHTMLTags(val)
   Dim re
   Set re=New RegExp
   re.IgnoreCase=True
   re.Pattern="<[^>]*>"
   re.Global=True
   stripHTMLTags=re.Replace(val,"")
   set re=Nothing
End Function

Function CleanHTML(val,acleancrlf)
   CleanHTML=stripHTMLTags(val)
   CleanHTML=replace(CleanHTML,"&lt;","<")
   CleanHTML=replace(CleanHTML,"&quot;","""")
   CleanHTML=replace(CleanHTML,"&gt;",">")
   CleanHTML=replace(CleanHTML,"&nbsp;"," ")
   CleanHTML=replace(CleanHTML,"&amp;","&")
   CleanHTML=replace(CleanHTML,"&bdquo;","„")
   CleanHTML=replace(CleanHTML,"&rdquo;","”")
   if acleancrlf=1 then
      CleanHTML=replace(CleanHTML,chr(10),"")
      CleanHTML=replace(CleanHTML,chr(13),"")
   end if
End Function

Function CleanHTMLEditor(val)
   CleanHTMLEditor=CleanHTML(val,0)
   CleanHTMLEditor=replace(CleanHTMLEditor,"&oacute;","ó")
   CleanHTMLEditor=replace(CleanHTMLEditor,"&Oacute;","Ó")
   CleanHTMLEditor=replace(CleanHTMLEditor,vbCrLf+vbCrLf,vbCrLf)
End Function

Function HTMLTextOnlyReplace(strInput, strPattern, strReplace) 
  ' Use <*> to indicate the match you wish to replace 
  Dim regEx, Match, Matches, Position, strReturn, sc1, sc2
  Position = 1 
  strReturn = "" 
  strPattern="(?![^<]+>)" + strPattern + "(?![^<]+>)"
  Set regEx = New RegExp 
  regEx.Pattern = strPattern 
  regEx.IgnoreCase = True 
  regEx.Global = True 
  Set Matches = regEx.Execute(strInput) 
  For Each Match in Matches 
    sc1=" "
    sc2=" "
    if Match.FirstIndex>0 then
      sc1=Mid(strInput,Match.FirstIndex,1)
    end if
    if (Match.FirstIndex+len(Match.Value))<Len(strInput) then
      sc2=Mid(strInput,Match.FirstIndex+len(Match.Value)+1,1)
    end if
    strReturn=strReturn & Mid(strInput, Position, Match.FirstIndex+1-Position)
    if (sc1>="a" and sc1<="z") or (sc1>="A" and sc1<="Z") or (sc2>="a" and sc2<="z") or (sc2>="A" and sc2<="Z") then
       strReturn=strReturn & Match.Value
    else
       strReturn=strReturn & Replace(strReplace, "<*>", Match.Value)
    end if
    Position = Len(Match.Value) + Match.FirstIndex + 1 
  Next 
  strReturn = strReturn & Mid(strInput, Position, Len(strInput))
  HTMLTextOnlyReplace = strReturn 
End Function 

Function GetTextAsCaption(s,amax)
   if len(s)>amax then
      GetTextAsCaption=Mid(s,1,amax)+"..."
   else
      GetTextAsCaption=s
   end if
End Function

sub RSCorrectTextFields(rs_tt,adefault)
   dim atv_tt,tt_field
   for each tt_field in rs_tt.fields
      if tt_field.type=201 then
         atv_tt=""
         if adefault=0 then
            atv_tt=tt_field.value
         end if
         tt_field.value=atv_tt
      end if
   next
end sub

sub RSAddNew(rs_tt)
   rs_tt.AddNew
end sub

sub RSUpdate(rs_tt)
   call RSCorrectTextFields(rs_tt,0)
   rs_tt.Update
end sub

Class EnhancedFormClass
 Public Function GetFormCollection()
  Dim FormFields  ' Dictionary which will store source fields.
  Dim aCType
  Set FormFields = Server.CreateObject("Scripting.Dictionary")
  FormFields.CompareMode=1

  'If there are some POST source data
  aCType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")
  if Request.ServerVariables("HTTP_CONTENT_TYPE")<>"" then
    If (Request.TotalBytes>0) And (aCType(0)="application/x-www-form-urlencoded") Then
      Dim SourceData
      SourceData = Request.BinaryRead(Request.TotalBytes)
      SourceData = RSBinaryToString(SourceData)   ' Convert source binary data to a string
  
      SourceData = Split(SourceData, "&")   ' Form fields are separated by "&"
      Dim Field, FieldName, FieldContents
      On Error Resume Next
      For Each Field In SourceData
        Field = Split(Field, "=")  ' Field name and contents is separated by "="
        FieldName = URL.decode(Field(0))
        if err.number<>0 then
           FieldName = URLDecode(Field(0))
        end if
        FieldContents = UTF2Win(URL.decode(Field(1)))
        if err.number<>0 then
           FieldContents = URLDecode(Field(1))
        end if
        FormFields.Add FieldName, FieldContents   ' Add field to the dictionary
      Next
      On Error Goto 0
    End If ' Request.Totalbytes > 0
  end if
  Set GetFormCollection = FormFields
 End Function
 
 Private Function DecodeUrlData(ByVal What)
  Dim Pos, pPos
  What = Replace(What, "+", " ") ' replace + to space

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
      Stream.WriteText Mid(What, pPos, Pos - pPos) + _
        Chr(CLng("&H" & Mid(What, Pos + 1, 2)))
      pPos = Pos + 3
      Pos = InStr(pPos, What, "%")
    Loop
    Stream.WriteText Mid(What, pPos)

    Stream.Position = 0 ' read the text stream
    DecodeUrlData = Stream.ReadText
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
    DecodeUrlData = What
  End If
 End Function
 
 Private Function RSBinaryToString(Binary)
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = Server.CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
 End Function

 Public Function AsString()
   On Error Resume Next
   AsString=""
   akeys=FormFields.Keys()
   For Each akey in akeys
      AsString=AsString & akey & "=" & server.URLEncode(left(FormFields(akey),1000)) & "&"
   Next
   On Error Goto 0
 End Function

 Public Function AsStringMaxValueLen(amaxlen)
   On Error Resume Next
   dim svalue,skey
   AsStringMaxValueLen=""
   akeys=FormFields.Keys()
   For Each akey in akeys
      svalue=ReplaceIllegalCharsLev0Text(left(FormFields(akey),amaxlen))
      skey=ReplaceIllegalCharsLev0Text(left(akey,amaxlen))
      if svalue<>"" and skey<>"" then
         AsStringMaxValueLen=AsStringMaxValueLen & URL.encodeSimple(skey) & "=" & URL.encodeSimple(svalue) & "&"
      end if
   Next
   On Error Goto 0
 End Function
 
 Public Function AsStringKeysOnly(amaxlen)
   On Error Resume Next
   dim svalue,skey
   AsStringKeysOnly=""
   akeys=FormFields.Keys()
   For Each akey in akeys
      svalue=ReplaceIllegalCharsLev0Text(left(FormFields(akey),amaxlen))
      skey=ReplaceIllegalCharsLev0Text(left(akey,amaxlen))
      if svalue<>"" and akey<>"" then
         AsStringKeysOnly=AsStringKeysOnly & URL.encodeSimple(skey) & chr(10)
      end if
   Next
   On Error Goto 0
 End Function
 
End Class

Function CheckCAPTCHA(valCAPTCHA)
   dim SessionCAPTCHA
   SessionCAPTCHA=Trim(Dession("CAPTCHA"))
   Dession("CAPTCHA")=vbNullString
   if Len(SessionCAPTCHA)<1 then
        CheckCAPTCHA=False
        exit function
   end if
   if CStr(SessionCAPTCHA)=UCase(CStr(valCAPTCHA)) then
       CheckCAPTCHA=True
   else
       CheckCAPTCHA=False
   end if
End Function

Function CheckCAPTCHAMulti(valCAPTCHA,valSufiks)
   dim SessionCAPTCHA
   SessionCAPTCHA=Trim(Dession("CAPTCHA"+valSufiks))
   Dession("CAPTCHA"+valSufiks)=vbNullString
   if Len(SessionCAPTCHA)<1 then
        CheckCAPTCHAMulti=False
        exit function
   end if
   if CStr(SessionCAPTCHA)=UCase(CStr(valCAPTCHA)) then
       CheckCAPTCHAMulti=True
   else
       CheckCAPTCHAMulti=False
   end if
End Function

Function CleanString(ByVal s, alen)
   Dim i, s1
   CleanString = ""
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanString = CleanString & s1
      End If
   Next
   If (alen > 0) And (Len(CleanString) <> alen) Then
      CleanString = ""
   End If
End Function

Function CleanStringBook(ByVal s, alen)
   Dim i, s1, s2
   CleanStringBook = ""
   if s="" then
      Exit Function
   end if
   s2 = ""
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Or s1 = "x" Or s1 = "X" Then
         s2 = s2 & s1
      End If
   Next
   if s2="" then
      Exit Function
   end if
   CleanStringBook = Mid(s2, Len(s2), 1)
   For i = Len(s2) - 1 To 1 Step -1
      s1 = Mid(s2, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanStringBook = s1 & CleanStringBook
      End If
   Next
   If (alen > 0) And (Len(CleanStringBook) <> alen) Then
      CleanStringBook = ""
   End If
End Function

Function CleanStringMusic(ByVal s, alen)
   Dim i, s1
   CleanStringMusic = ""
   if s="" then
      Exit Function
   end if
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (CleanStringMusic = "") And (s1 = "m" Or s1 = "M") Then
         CleanStringMusic = CleanStringMusic & "3"
      ElseIf (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanStringMusic = CleanStringMusic & s1
      End If
   Next
   If (alen > 0) And (Len(CleanStringMusic) <> alen) Then
      CleanStringMusic = ""
   End If
End Function

Function ValidateREGON9(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateREGON9 = 0
   wagi = Array(8, 9, 2, 3, 4, 5, 6, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (suma Mod 11) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateREGON9 = 1
   End If
End Function

Function ValidateREGON14(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateREGON14 = 0
   wagi = Array(2, 4, 8, 5, 0, 9, 7, 3, 6, 1, 2, 4, 8)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   If ValidateREGON9(Mid(s, 1, 9)) = 0 Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (suma Mod 11) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateREGON14 = 1
   End If
End Function

Function ValidateREGON(anumer)
   Dim s
   s = CleanString(anumer, 0)
   ValidateREGON = 0
   If Len(s) = 9 Then
      ValidateREGON = ValidateREGON9(s)
   ElseIf Len(s) = 14 Then
      ValidateREGON = ValidateREGON14(s)
   End If
End Function

Function ValidatePESEL(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidatePESEL = 0
   wagi = Array(1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidatePESEL = 1
   End If
End Function

Function ValidateNIP(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateNIP = 0
   wagi = Array(6, 5, 7, 2, 3, 4, 5, 6, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = suma Mod 11
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateNIP = 1
   End If
End Function

Function ValidateISBN(anumer)
   Dim s, i, wagi, anum, suma, aarr, acheck
   ValidateISBN = 0
   wagi = Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
   s = CleanStringBook(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (11 - (suma Mod 11)) Mod 11
   If Mid(s, i, 1) = "x" Or Mid(s, i, 1) = "X" Then
      acheck = 10
   Else
      acheck = CLng(Mid(s, i, 1))
   End If
   If suma = acheck Then
      ValidateISBN = 1
   End If
End Function

Function ValidateISSN(anumer)
   Dim s, i, wagi, anum, suma, aarr, acheck
   ValidateISSN = 0
   wagi = Array(8, 7, 6, 5, 4, 3, 2)
   s = CleanStringBook(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (11 - (suma Mod 11)) Mod 11
   If Mid(s, i, 1) = "x" Or Mid(s, i, 1) = "X" Then
      acheck = 10
   Else
      acheck = CLng(Mid(s, i, 1))
   End If
   If suma = acheck Then
      ValidateISSN = 1
   End If
End Function

Function ValidateISMN(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateISMN = 0
   wagi = Array(3, 1, 3, 1, 3, 1, 3, 1, 3)
   s = CleanStringMusic(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateISMN = 1
   End If
End Function

Function ValidateEAN13(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateEAN13 = 0
   wagi = Array(1, 3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateEAN13 = 1
   End If
End Function

Function ValidateCPV(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateCPV = 0
   wagi = Array(3, 7, 1, 3, 7, 1, 3, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = suma Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateCPV = 1
   End If
End Function

Function ValidateEAN8(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateEAN8 = 0
   wagi = Array(3, 1, 3, 1, 3, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateEAN8 = 1
   End If
End Function

Function ValidateFloat(strValue)
   on error resume next
   ValidateFloat=1
   If InStr(strValue,".")=0 AND NOT IsNumeric(strValue) Then
      ValidateFloat=0
   End If
   on error goto 0
End Function

Function ValidateInteger(strValue)
   on error resume next
   ValidateInteger=1
   If NOT IsNumeric(strValue) OR NOT CDbl(strValue)=CLng(strValue) Then
      ValidateInteger=0
   End If
   on error goto 0
End Function

Function ValidateDate(strValue)
   on error resume next
   ValidateDate=1
   If NOT IsDate(strValue) Then
      ValidateDate=0
   End If
   on error goto 0
End Function

function getDateTimeAsStr(sdatetime)
   getDateTimeAsStr=""
   if year(sdatetime)<>1900 then
      getDateTimeAsStr=CStr(year(sdatetime)) & "/" & Right("0" & CStr(Month(sdatetime)),2) & "/" & Right("0" & CStr(Day(sdatetime)),2)
   end if
   if (hour(sdatetime)<>0) or (minute(sdatetime)<>0) then
      if getDateTimeAsStr<>"" then
         getDateTimeAsStr=getDateTimeAsStr & " "
      end if
      getDateTimeAsStr=getDateTimeAsStr & Right("0" & CStr(hour(sdatetime)),2) & ":" & Right("0" & CStr(minute(sdatetime)),2)
   end if
end function


function getDateAsSQLDate(sdatetime)
   dim i,sarr,sdarr,sdate,stime
   sdatetime=Replace(Replace(Replace(trim(sdatetime),".","/"),"-","/"),"","/")
   if Not IsDate(sdatetime) or len(sdatetime)<8 then
      getDateAsSQLDate=""
   else
      sarr=Split(sdatetime," ")
      for i=0 to ubound(sarr)
         if instr(sarr(i),":")>0 then
            stime=" " & sarr(i)
         elseif sarr(i)<>"" then
            sdarr=Split(sarr(i),"/")
            if CLng(sdarr(0))>31 then
               sdate=sdarr(0) & Right("0" & sdarr(1),2) & Right("0" & sdarr(2),2)
            else
               sdate=sdarr(2) & Right("0" & sdarr(1),2) & Right("0" & sdarr(0),2)
            end if
         end if
      next               
      getDateAsSQLDate=trim(sdate & stime)
   end if
end function

function getStrAsDateTime(sdate,shour,smin,anopastdate)
   dim aa,bb
   if Not IsDate(sdate) then
      getStrAsDateTime=DateSerial(1900,1,1)+TimeSerial(0,0,0)
   else
      bb=Split(sdate," ",-1,1)
      aa=Split(Replace(Replace(Replace(trim(CStr(bb(0))),".","/"),"-","/"),"","/"),"/",-1,1)
      getStrAsDateTime=DateSerial(CLng(aa(0)),CLng(aa(1)),CLng(aa(2)))
      if anopastdate=1 then
         if getStrAsDateTime<Now then
            getStrAsDateTime=Now
            Exit Function
         end if
      end if
      if (shour<>"") and (smin<>"") then
         getStrAsDateTime=getStrAsDateTime+TimeSerial(CLng(shour),CLng(smin),0)
      end if
   end if
end function

Function IIf(ByVal Cn, ByVal T, ByVal F)
  If Cn Then
    IIf = T
  Else
    IIf = F
  End If
End Function

Function IsN(ByVal str)
  isN = False
  Select Case VarType(str)
    Case vbEmpty, vbNull
      isN = True : Exit Function
    Case vbString
      If str="" Then isN = True : Exit Function
    Case vbObject
      If TypeName(str)="Nothing" Or TypeName(str)="Empty" Then isN = True : Exit Function
    Case vbArray,8194,8204,8209
      If Ubound(str)=-1 Then isN = True : Exit Function
    End Select
End Function

Function GetIP()
   Dim addr, x, y
   x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
   y = Request.ServerVariables("REMOTE_ADDR")
   addr = IIF(IsN(x) or lCase(x)="unknown",y,x)
   If InStr(addr,".")=0 Then 
      addr = "0.0.0.0"
   End If
   GetIP = addr
End Function

Sub SetApp(AppName,AppData)
   Application.Lock
   Application.Contents.Item(AppName) = AppData
   Application.UnLock
End Sub

Function GetApp(AppName)
   If IsN(AppName) Then 
      GetApp = ""
      Exit Function
   End If
   GetApp = Application.Contents.Item(AppName)
End Function

Sub RemoveApp(AppName)
   Application.Lock
   Application.Contents.Remove(AppName)
   Application.UnLock
End Sub

Function JsEncode(ByVal str)
  If Not IsN(str) Then
    str = Replace(str,Chr(92),"\\")
    str = Replace(str,Chr(34),"\""")
    str = Replace(str,Chr(39),"\'")
    str = Replace(str,Chr(9),"\t")
    str = Replace(str,Chr(13),"\r")
    str = Replace(str,Chr(10),"\n")
    str = Replace(str,Chr(12),"\f")
    str = Replace(str,Chr(8),"\b")
  End If
  JsEncode = str
End Function

Function URLEscape(ByVal str)
  Dim i,c,a,s,s1
  s = ""
  If IsN(str) Then URLEscape = "" : Exit Function
  For i = 1 To Len(str)
    c = Mid(str,i,1)
    a = ASCW(c)
    If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
      s = s & c
    ElseIf InStr("@*_+-./:~?#",c)>0 Then
      s = s & c
    ElseIf a>0 and a<16 Then
      s = s & "%0" & Hex(a)
    ElseIf a>=16 and a<256 Then
      s = s & "%" & Hex(a)
    Else
      s1=Hex(a)
      if len(s1)<4 then
         s1="0" & s1
      end if
      if len(s1)<4 then
         s1="0" & s1
      end if
      s = s & "%u" & s1
    End If
  Next
  URLEscape = s
End Function 

Function URLUnEscape(ByVal str)
  Dim x, s
  x = InStr(str,"%")
  s = ""
  Do While x>0
    s = s & Mid(str,1,x-1)
    If LCase(Mid(str,x+1,1))="u" Then
      s = s & ChrW(CLng("&H"&Mid(str,x+2,4)))
      str = Mid(str,x+6)
    Else
      s = s & Chr(CLng("&H"&Mid(str,x+1,2)))
      str = Mid(str,x+3)
    End If
    x=InStr(str,"%")
  Loop
  URLUnEscape = s & str
End Function

Function RegexpTest(ByVal Str, ByVal Pattern)
  If IsN(Str) Then RegexpTest = False : Exit Function
  Dim Reg
  Set Reg = New RegExp
  Reg.IgnoreCase = True
  Reg.Global = True
  Reg.Pattern = Pattern
  RegexpTest = Reg.Test(CStr(Str))
  Set Reg = Nothing
End Function

Function RegexpReplace(ByVal Str, ByVal rule, Byval Result, ByVal isM)
  Dim tmpStr,Reg : tmpStr = Str
  If Not IsN(Str) Then
    Set Reg = New Regexp
    Reg.Global = True
    Reg.IgnoreCase = True
    If isM = 1 Then Reg.Multiline = True
    Reg.Pattern = rule
    tmpStr = Reg.Replace(tmpStr,Result)
    Set Reg = Nothing
  End If
  RegexpReplace = tmpStr
End Function

Function RegexpMatch(ByVal Str, ByVal rule)
  Dim Reg
  Set Reg = New Regexp
  Reg.Global = True
  Reg.IgnoreCase = True
  Reg.Pattern = rule
  Set RegexpMatch = Reg.Execute(Str)
  Set Reg = Nothing
End Function

' Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
' as set out in the memo RFC1321.
'
'

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)

    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)

    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

Private Function md5_f(x, y, z)
    md5_f = (x And y) Or ((Not x) And z)
End Function

Private Function md5_G(x, y, z)
    md5_G = (x And z) Or (y And (Not z))
End Function

Private Function md5_H(x, y, z)
    md5_H = (x Xor y Xor z)
End Function

Private Function md5_i(x, y, z)
    md5_i = (y Xor (x Or (Not z)))
End Function

Private Sub md5_FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_f(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub md5_II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_i(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount

    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448

    lMessageLength = Len(sMessage)

    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)

    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount

    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Public Function MD5(sMessage)
    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d

    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    x = ConvertToWordArray(sMessage)

    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d

        md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
        md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
        md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
        md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
        md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
        md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
        md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
        md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
        md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
        md5_FF b, c, d, a, x(k + 15), S14, &H49B40821

        md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
        md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
        md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
        md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
        md5_GG d, a, b, c, x(k + 10), S22, &H2441453
        md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
        md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
        md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
        md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A

        md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
        md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
        md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665

        md5_II a, b, c, d, x(k + 0), S41, &HF4292244
        md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
        md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
        md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
        md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
        md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
        md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        md5_II c, d, a, b, x(k + 6), S43, &HA3014314
        md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
        md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
        md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
        md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        md5_II b, c, d, a, x(k + 9), S44, &HEB86D391

        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next

    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function

Sub arrayAdd(byref aarr,byref aval)
   ReDim Preserve aarr(UBound(aarr) + 1)
   if isobject(aval) then
      set aarr(UBound(aarr)) = aval
   else
      aarr(UBound(aarr)) = aval
   end if
End Sub

%>