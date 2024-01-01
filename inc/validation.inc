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
    strExtensions = Split("com|net|edu|eu|fm|gov|int|net|org|mil|info|cc|biz|name|pro|at|au|be|bg|br|by|ca|ch|cn|cy|cz|de|dk|ee|es|fi|fr|fx|gb|gr|hr|hu|ie|il|it|li|lt|lv|md|mt|mx|nl|no|nz|pe|pl|pt|ro|ru|se|si|sk|tr|tv|ua|uk|us", "|") 'kompatybilne z /icorlib/validation/email_check.js

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

Function UTF2Win(s)
   s=replace(s,chr(196) & chr(133),chr(185))
   s=replace(s,chr(196) & chr(135),chr(230))
   s=replace(s,chr(196) & chr(153),chr(234))
   s=replace(s,chr(197) & chr(130),chr(179))
   s=replace(s,chr(197) & chr(132),chr(241))
   s=replace(s,chr(195) & chr(179),chr(243))
   s=replace(s,chr(197) & chr(155),chr(156))
   s=replace(s,chr(197) & chr(186),chr(159))
   s=replace(s,chr(197) & chr(188),chr(191))
   s=replace(s,chr(196) & chr(132),chr(165))
   s=replace(s,chr(196) & chr(134),chr(198))
   s=replace(s,chr(196) & chr(152),chr(202))
   s=replace(s,chr(197) & chr(129),chr(163))
   s=replace(s,chr(197) & chr(131),chr(209))
   s=replace(s,chr(195) & chr(147),chr(211))
   s=replace(s,chr(197) & chr(154),chr(140))
   s=replace(s,chr(197) & chr(185),chr(143))
   s=replace(s,chr(197) & chr(187),chr(175))
   UTF2Win=s
End Function


Function RegExReplace(strString,strPattern,strReplace)
   Dim oRegex
   Set oRegex = New RegExp
   oRegex.IgnoreCase = True
   oRegex.Global=True
   oRegex.Pattern=strPattern
   If oRegex.Test(strString) Then
      RegExReplace = oRegex.Replace(strString, strReplace)
   Else
      RegExReplace=strString
   End If
   Set oRegex = Nothing
End Function

Function UTF2NoPL(s)
   s=RegExReplace(s,"("+chr(196)+chr(133)+")", "a")
   s=RegExReplace(s,"("+chr(196)+chr(135)+")", "c")
   s=RegExReplace(s,"("+chr(196)+chr(153)+")", "e")
   s=RegExReplace(s,"("+chr(197)+chr(130)+")", "l")
   s=RegExReplace(s,"("+chr(197)+chr(132)+")", "n")
   s=RegExReplace(s,"("+chr(195)+chr(179)+")", "o")
   s=RegExReplace(s,"("+chr(197)+chr(155)+")", "s")
   s=RegExReplace(s,"("+chr(197)+chr(188)+"|"+chr(197)+chr(186)+")", "z")   
   s=RegExReplace(s,"("+chr(196)+chr(132)+")", "A")
   s=RegExReplace(s,"("+chr(196)+chr(134)+")", "C")
   s=RegExReplace(s,"("+chr(196)+chr(152)+")", "E")
   s=RegExReplace(s,"("+chr(197)+chr(129)+")", "L")
   s=RegExReplace(s,"("+chr(197)+chr(131)+")", "N")
   s=RegExReplace(s,"("+chr(195)+chr(147)+")", "O")
   s=RegExReplace(s,"("+chr(197)+chr(154)+")", "S")
   s=RegExReplace(s,"("+chr(197)+chr(187)+"|"+chr(197)+chr(185)+")", "Z")   
   
   s=RegExReplace(s,"(\u0105)", "a")
   s=RegExReplace(s,"(\u0107)", "c")
   s=RegExReplace(s,"(\u0119)", "e")
   s=RegExReplace(s,"(\u0142)", "l")
   s=RegExReplace(s,"(\u0144)", "n")
   s=RegExReplace(s,"(\u00F3)", "o")
   s=RegExReplace(s,"(\u015B)", "s")
   s=RegExReplace(s,"(\u017A|\u017C)", "z")   
   s=RegExReplace(s,"(\u0104)", "A")
   s=RegExReplace(s,"(\u0106)", "C")
   s=RegExReplace(s,"(\u0118)", "E")
   s=RegExReplace(s,"(\u0141)", "L")
   s=RegExReplace(s,"(\u0143)", "N")
   s=RegExReplace(s,"(\u00D3)", "O")
   s=RegExReplace(s,"(\u015A)", "S")
   s=RegExReplace(s,"(\u0179|\u017B)", "Z")
   UTF2NoPL=s
End Function

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
   dim re
   set re = new RegExp
   re.pattern = "<[\\w/]+[^<>]*>"
   re.global=true
   stripHTMLTags = re.replace(val,"")
   set re=Nothing
End Function

Function CleanHTML(val,acleancrlf)
   CleanHTML=stripHTMLTags(val)
   CleanHTML=replace(CleanHTML,"&lt;","<")
   CleanHTML=replace(CleanHTML,"&quot;","""")
   CleanHTML=replace(CleanHTML,"&gt;",">")
   CleanHTML=replace(CleanHTML,"&nbsp;"," ")
   CleanHTML=replace(CleanHTML,"&amp;","&")
   if acleancrlf=1 then
      CleanHTML=replace(CleanHTML,chr(10),"")
      CleanHTML=replace(CleanHTML,chr(13),"")
   end if
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

Sub arrayAdd(byref aarr,byref aval)
   ReDim Preserve aarr(UBound(aarr) + 1)
   if isobject(aval) then
      set aarr(UBound(aarr)) = aval
   else
      aarr(UBound(aarr)) = aval
   end if
End Sub
%>
