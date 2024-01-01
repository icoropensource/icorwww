<%
Function ReplaceIllegalCharsLev0(sInput) 
   Dim sBadChars, iCounter
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "_", "#", "%", "&", "'", "(", ")", ":", ";", "<", ">", "=", "[", "]", "?", "`", "|", "/", "\" )
   For iCounter=0 to uBound(sBadChars) 
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next 
   ReplaceIllegalCharsLev0=sInput
End function

Function ReplaceIllegalCharsLev1(sInput) 
   Dim sBadChars, iCounter
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "_", "#", "%", "&", "'", ":", ";", "?", "`", "|" )
   For iCounter=0 to uBound(sBadChars) 
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next 
   ReplaceIllegalCharsLev1=sInput
End function

Function ReplaceIllegalCharsLev2(sInput) 
   Dim sBadChars, iCounter
   sBadChars=array("select", "drop", "--", "insert", "delete", "xp_", "truncate", "update", _
      "#", "%", "&", "'", "`", "|" )
   For iCounter=0 to uBound(sBadChars) 
      sInput=replace(sInput,sBadChars(iCounter),"")
   Next 
   ReplaceIllegalCharsLev2=sInput
End function

function getStringDateAsString(sdatetime)
   dim sdarr,ayear,amonth,aday
   sdatetime=Replace(Replace(Replace(trim(sdatetime),".","/"),"-","/"),"","/")
   getStringDateAsString=""
   if IsDate(sdatetime) then
      sdarr=Split(sdatetime,"/")
      if UBound(sdarr)=2 then
         ayear=sdarr(0)
         amonth=sdarr(1)
         aday=sdarr(2)
         if CInt(ayear)<1700 or CInt(ayear)>2030 then
            Exit Function
         end if
         if CInt(amonth)<1 or CInt(amonth)>12 then
            Exit Function
         end if
         if CInt(aday)<1 or CInt(aday)>31 then
            Exit Function
         end if
         getStringDateAsString=trim(ayear & "-" & Right("0" & amonth,2) & "-" & Right("0" & aday,2))
      end if
   end if
end function

function getSQLDateAsString(bdate)
   getSQLDateAsString=CStr(bdate)
   if mid(getSQLDateAsString,1,10)="1900-01-01" then
      getSQLDateAsString=""
   end if
end function

%>