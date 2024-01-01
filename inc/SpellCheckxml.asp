<%@ CODEPAGE=65001 LANGUAGE="VBSCRIPT" %>
<%
Response.AddHeader "Pragma", "no-cache"
Response.expires=0
Response.ContentType="text/xml"
Response.Charset ="utf-8"
Response.Buffer=true
on error resume next
tStart=time
strAspellPath=""
strTempPath=Application("DefaultTempPath")
Set WshShell = Server.CreateObject( "WScript.Shell" )
strAspellPath=WSHShell.RegRead("HKLM\Software\Aspell\")
set xmlDoc=server.CreateObject("MSXML2.DOMDocument")
xmlDoc.async=false
ret=xmlDoc.load(Request)
function Translate(s,par)
   dim t1,t2,i
   t1=Array(185,156,159,165,140,143)
   t2=Array(177,182,188,161,166,172)
   for i=0 to 5
      if par=1 then
        'win1250->iso88592
        s=replace(s,chr(t1(i)),chr(t2(i)))
      else
        'iso88592->win1250
        s=replace(s,chr(t2(i)),chr(t1(i)))
      end if
   next
   Translate=s
end function
if ret and strAspellPath<>"" then
   set adocelement=xmlDoc.documentElement
   amode=adocelement.getAttribute("mode")   
   alanguage=adocelement.getAttribute("language")
   atext=adocelement.firstChild.text
   if amode="check" then
      Set objFSO = CreateObject("Scripting.FileSystemObject")
      strFileName = objFSO.GetTempName
      strFullName = objFSO.BuildPath(strTempPath, strFileName)
      if objFSO.FileExists(strFullName) then
         objFSO.DeleteFile strFullName,True
      end if
      Set objFile = objFSO.CreateTextFile(strFullName)
      atext=replace(atext,vbcr,"")
      atext=Translate(atext,1)
      lines=split(atext,vblf)
      nlines=ubound(lines)+1
      objFile.Write "^" & replace(atext,vblf,vblf & "^")
      objFile.Close   
      if len(atext)<1000 then
        aspellspeed="normal"
      else  
        aspellspeed="fast"
      end if
      aspellspeed="fast"
      cmd="%comspec% /c """ & strAspellPath & "\bin\aspell.exe"" -a --ignore-case --sug-mode=" & aspellspeed & "  --lang=" & alanguage & " <" & strFullName
      set Pipe = WshShell.Exec(cmd)
      if Not Pipe.StdOut.AtEndOfStream then
         line=Pipe.StdOut.ReadLine
         if left(line,4)="@(#)" then
            l=1
            w=1
            m=1
            wordoffset=0
            amistakes=""
            Response.Write "<?xml version=""1.0"" encoding=""utf-16""?>" & vbcrlf
            Response.Write "<SPELLCHECK>" & vbcrlf
            xml_len=160000
            isoverflowed="0"
            Do While Not Pipe.StdOut.AtEndOfStream
               line=Translate(Pipe.StdOut.ReadLine,2)
               flag=left(line,1) 
               line=mid(line,3)
               if flag<>"" then         
                  wordoffset=wordoffset+1
                  if flag="#" or flag="&" then
                     if flag="#" then
                         head=split(line)
                        original=head(0)
                        offset=head(1)
                        sugcount=0
                        sug=""
                     elseif flag="&" then
                        pom_arr=split(line,":")
                        head=split(pom_arr(0))
                        original=head(0)
                        sugcount=head(1)
                        offset=head(2)
                        sugestions=split(pom_arr(1),", ")
                        sug=pom_arr(1)
                     end if
                     pom="<M w=""" & wordoffset & """ o=""" & original & """ f=""" & offset-1 & """>"
                     if sugcount>8 then
                        sugcount=8
                     end if   
                     for i=1 to sugcount
                        pom=pom & "<S>" & trim(sugestions(i-1)) & "</S>"
                     next
                     pom=pom & "</M>"
                     if w=1 then
                        amistakes=pom 
                     else
                        amistakes=amistakes & pom
                     end if
                     m=m+1
                  end if
                  w=w+1
               else
                  alinestr="<L l=""" & len(lines(l-1)) & """>"
                  if xml_len>0 then
                     Response.Write alinestr
                     Response.Write amistakes
                     Response.Write "</L>"
                  else
                     isoverflowed="1"
                     Pipe.StdOut.ReadAll
                  end if
                  l=l+1
                  w=1
                  xml_len=xml_len-len(alinestr)-len(amistakes)-8
                  amistakes=""
               end if
            Loop
            tEnd=time
            
            Response.Write "<STATS lines=""" & nlines & """ wordcount=""" & wordoffset & """ isoverflowed=""" & isoverflowed & """ time=""" & DateDiff("s", tStart , tEnd) & """ />"  & vbcrlf
            Response.Write "</SPELLCHECK>" & vbcrlf
         end if
      end if
      if objFSO.FileExists(strFullName) then
         objFSO.DeleteFile strFullName,True
      end if
      Set objFile=Nothing
      Set objFSO=Nothing      
   elseif amode="add" then
      cmd="%comspec% /c """ & strAspellPath & "\bin\aspell.exe"" -a --lang=" & alanguage
      set Pipe = WshShell.Exec(cmd)   
      Pipe.StdIn.WriteLine "*" & Translate(atext,1) & vbcrlf & "#" & chr(26)
   end if
   set Pipe=nothing
elseif strAspellPath="" then
   Response.Write "NOT AVAILABLE"
else
   Response.Write "ERROR"
end if
Set WshShell=nothing
set xmlDoc=Nothing
if err.number<>0 then
   set xmlDoc=Nothing
   Set WshShell=Nothing
   Set objFile=Nothing
   Set objFSO=Nothing
end if
on error goto 0
%>
