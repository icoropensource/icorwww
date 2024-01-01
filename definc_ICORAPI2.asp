<%
ICOR_TCP_SERVER=Application("ICOR_TCP_SERVER2")
ICOR_TCP_SERVER_PORT=Application("ICOR_TCP_SERVER_PORT2")
ICOR_TCP_SERVER_PAGE=Application("ICOR_TCP_SERVER_PAGE2")

'Function _URLEncode(strData)
'   Dim i,c,ln
'   ln=Len(strData)
'   For i=1 To ln
'      c=Asc(Mid(strData,i,1))
'      If c=32 Then
'         _URLEncode=_URLEncode & "+"
'      ElseIf (c>=48 and c<=57) Or (c>=65 and c<=90) Or (c>=97 and c<=122) Or c=61 Or c=38 Then
'         _URLEncode=_URLEncode & Chr(c)
'      Else
'         c=Hex(CLng(c))
'         If Len(c)<2 Then c="0" & c
'         _URLEncode=_URLEncode & "%" & c
'      End If
'   Next
'End Function

Function PadHex(anumber)
   PadHex=Hex(CLng(anumber))
   PadHex=String(8-len(PadHex),"0")+PadHex
End Function

Function str2arr(s)
   alen=Len(s)\8
   Redim b(alen)
   For i=0 To alen-1
      b(i)=CLng("&H"+Mid(s,1+i*8,8))
   Next
   str2arr=b
End Function

Function utf8Len( text )
    utf8Len = 0
    ' Only strings with data
    If VarType( text ) <> 8 Then Exit Function
    If Len( text ) < 1 Then Exit Function
    ' Create an ADODB.Stream object to handle charset conversion
    With (CreateObject("ADODB.Stream"))
        ' Define characteristics of the stream
        .Type = 2 '( adTypeText )
        .Charset = "utf-8"
        .Open
        .WriteText text
        .Flush
        ' Get the length of the stream without BOM
        utf8Len = .Size - 3
        .Close
    End With 
End Function

Function RSBinaryToString(xBinary)
   Dim Binary
   If VarType(xBinary)=8 Then Binary=MultiByteToBinary(xBinary) Else Binary=xBinary
      Dim RS, LBinary
      Const adLongVarChar=201
      Set RS=CreateObject("ADODB.Recordset")
      LBinary=LenB(Binary)
      If LBinary>0 Then
      RS.Fields.Append "mBinary",adLongVarChar,LBinary
      RS.Open
      RS.AddNew
      RS("mBinary").AppendChunk Binary 
      RS.Update
      RSBinaryToString=RS("mBinary")
      Set RS=Nothing
   Else
      RSBinaryToString=""
   End If
   Set RS=Nothing
End Function

Function MultiByteToBinary(MultiByte)
   Dim RS, LMultiByte, Binary
   Const adLongVarBinary = 205
   Set RS=CreateObject("ADODB.Recordset")
   LMultiByte=LenB(MultiByte)
   If LMultiByte>0 Then
      RS.Fields.Append "mBinary",adLongVarBinary,LMultiByte
      RS.Open
      RS.AddNew
      RS("mBinary").AppendChunk MultiByte & ChrB(0)
      RS.Update
      Binary=RS("mBinary").GetChunk(LMultiByte)
   End If
   Set RS=Nothing
   MultiByteToBinary=Binary
End Function

Function ICORCommand(acommand,adata)
   Dim xmlhttp,aicor
   ICORCommand=""
   adata=Server.URLEncode(adata)
   ldata=Len(adata)
   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
'   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.4.0")
   xmlhttp.setTimeouts 0, 120*1000, 15*60*1000, 30*60*1000
   addr="http://"+ICOR_TCP_SERVER+":"+ICOR_TCP_SERVER_PORT+"/"+acommand
   xmlhttp.Open "POST", addr, false
   xmlhttp.setRequestHeader "Content-Type", "application/icor;charset=utf-8" '"application/x-www-form-urlencoded"
   on error resume next
   xmlhttp.Send bdata
   if xmlhttp.status=200 Then
      ICORCommand=RSBinaryToString(xmlhttp.ResponseBody)
   else
      ICORCommand="could not communicate to ICOR Server (a): "+CStr(xmlhttp.status) + " : " + addr
      Exit Function
   end if
   set xmlhttp=Nothing
   on error goto 0
End Function

Function ICORMessage(adata)
   Dim xmlhttp,aicor
   ICORMessage=""
   adata=Server.URLEncode(adata)
   ldata=Len(adata)
   MAX_SIZE=60000000
   apackagecount=ldata\MAX_SIZE
   alastpackagesize=ldata-(apackagecount*MAX_SIZE)
'   Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
'   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
'   Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.4.0")
   xmlhttp.setTimeouts 0, 120*1000, 15*60*1000, 30*60*1000
   apos=1
   alast=0
   aid="00000000"
   do while apackagecount>0 or alastpackagesize>0
      if apackagecount>0 then
         asize=MAX_SIZE
         apackagecount=apackagecount-1
         if (apackagecount=0) and (alastpackagesize=0) then
            alast=1
         end if
      else
         asize=alastpackagesize
         alastpackagesize=0
         alast=1
      end if
      addr="http://"+ICOR_TCP_SERVER+":"+ICOR_TCP_SERVER_PORT+ICOR_TCP_SERVER_PAGE+"?msglen="+CStr(asize)+"&size="+CStr(ldata)+"&id="+aid
      xmlhttp.Open "POST", addr, false
      xmlhttp.setRequestHeader "Content-Type", "application/icor;charset=utf-8" '"application/x-www-form-urlencoded"
      bdata=mid(adata,apos,asize)
      on error resume next
      asendcnt=5
      do while asendcnt>0
         xmlhttp.Send bdata
         if err.number=0 then
            exit do
         end if
         if err.number=-2147012867 then
            exit do
         end if
         asendcnt=asendcnt-1
      loop
      if xmlhttp.status=200 Then
         if alast=1 then
            ICORMessage=RSBinaryToString(xmlhttp.ResponseBody)
         elseif apos=1 then
            aid=xmlhttp.ResponseText
         end if
      else
         ICORMessage="could not communicate to ICOR Server (b): "+CStr(xmlhttp.status) + " : " + addr
         Exit Do
      end if
      apos=apos+asize
   loop
   set xmlhttp=Nothing
   on error goto 0
End Function

Function AddObject(fparamI1,fparamI2)
   AddObject=CLng("&H"+ICORMessage(PadHex(0)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function CheckObjectBySummary(fparamI1,fparamI2,fparamI3,fparamI4,fparamI5)
   CheckObjectBySummary=CLng("&H"+ICORMessage(PadHex(1)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)))
End Function

Function ClassExists(fparamI1,fparamI2)
   ClassExists=CLng("&H"+ICORMessage(PadHex(2)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function ClearAllObjects(fparamI1,fparamI2)
   ClearAllObjects=CLng("&H"+ICORMessage(PadHex(3)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function ClearAllValues(fparamI1,fparamI2,fparamS1)
   ClearAllValues=CLng("&H"+ICORMessage(PadHex(4)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function ClearStdErr()
   ICORMessage(PadHex(5))
End Function

Function ClearStdOut()
   ICORMessage(PadHex(6))
End Function

Function CompareOIDValue(fparamI1,fparamI2,fparamS1,fparamI3,fparamS2)
   CompareOIDValue=CLng("&H"+ICORMessage(PadHex(7)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(len(fparamS2))+fparamS2))
End Function

Function CompareOIDs(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4)
   CompareOIDs=CLng("&H"+ICORMessage(PadHex(8)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)))
End Function

Function CreateObjectByID(fparamI1,fparamI2,fparamI3)
   ICORMessage(PadHex(9)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3))
End Function

Function DeleteObject(fparamI1,fparamI2,fparamI3)
   DeleteObject=CLng("&H"+ICORMessage(PadHex(10)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function DeleteVariable(fparamI1,fparamS1)
   ICORMessage(PadHex(11)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1)
End Function

Function DialogInput(fparamI1,fparamS1,fparamS2,fparamS3)
   DialogInput=ICORMessage(PadHex(12)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3)
End Function

Function DoEvents()
   ICORMessage(PadHex(13))
End Function

Function EditObject(fparamI1,fparamI2,fparamI3,fparamS1,fparamI4)
   EditObject=CLng("&H"+ICORMessage(PadHex(14)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI4)))
End Function

Function ExcelClose()
   ExcelClose=ICORMessage(PadHex(15))
End Function

Function ExcelExecute(fparamS1,fparamS2)
   ExcelExecute=ICORMessage(PadHex(16)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function ExcelGetCellValue(fparamI1,fparamI2)
   ExcelGetCellValue=ICORMessage(PadHex(17)+PadHex(fparamI1)+PadHex(fparamI2))
End Function

Function ExcelOpen(fparamI1)
   ExcelOpen=ICORMessage(PadHex(18)+PadHex(fparamI1))
End Function

Function ExcelSetCellValue(fparamI1,fparamI2,fparamS1)
   ExcelSetCellValue=ICORMessage(PadHex(19)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ExecuteMethod(fparamI1,fparamI2,fparamS1,fparamS2,fparamI3,fparamS3,fparamI4)
   'response.write len(fparamS2)
   'response.write "/"
   'response.write lenb(fparamS2)
   'response.end
   ExecuteMethod=ICORMessage(PadHex(20)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(utf8Len(fparamS1))+fparamS1+PadHex(utf8Len(fparamS2))+fparamS2+PadHex(fparamI3)+PadHex(utf8Len(fparamS3))+fparamS3+PadHex(fparamI4))
End Function

Function ExportModuleAsString(fparamI1,fparamS1,fparamS2)
   ExportModuleAsString=ICORMessage(PadHex(21)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function FindValue(fparamI1,fparamI2,fparamS1,fparamS2)
   FindValue=CLng("&H"+ICORMessage(PadHex(22)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2))
End Function

Function FindValueBoolean(fparamI1,fparamI2,fparamS1,fparamI3)
   FindValueBoolean=CLng("&H"+ICORMessage(PadHex(23)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function FindValueDateTime(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9)
   FindValueDateTime=CLng("&H"+ICORMessage(PadHex(24)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9)))
End Function

Function FindValueFloat(fparamI1,fparamI2,fparamS1,fparamD1)
   FindValueFloat=CLng("&H"+ICORMessage(PadHex(25)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(CStr(fparamD1)))+CStr(fparamD1)))
End Function

Function FindValueInteger(fparamI1,fparamI2,fparamS1,fparamI3)
   FindValueInteger=CLng("&H"+ICORMessage(PadHex(26)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function FormatFNum(fparamS1,fparamD1)
   FormatFNum=ICORMessage(PadHex(27)+PadHex(len(fparamS1))+fparamS1+PadHex(len(CStr(fparamD1)))+CStr(fparamD1))
End Function

Function GetClassID(fparamI1,fparamS1)
   GetClassID=CLng("&H"+ICORMessage(PadHex(28)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetClassLastModification(fparamI1,fparamI2)
   GetClassLastModification=str2arr(ICORMessage(PadHex(29)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function GetClassProperty(fparamI1,fparamI2,fparamS1)
   GetClassProperty=ICORMessage(PadHex(30)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1)
End Function

Function GetDeletedObjectsList(fparamI1,fparamI2,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9)
   GetDeletedObjectsList=ICORMessage(PadHex(32)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9))
End Function

Function GetFieldLastModification(fparamI1,fparamI2,fparamS1)
   GetFieldLastModification=str2arr(ICORMessage(PadHex(33)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetFieldModification(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldModification=ICORMessage(PadHex(34)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3))
End Function

Function GetFieldObjectsCount(fparamI1,fparamI2,fparamS1)
   GetFieldObjectsCount=CLng("&H"+ICORMessage(PadHex(35)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetFieldProperty(fparamI1,fparamI2,fparamS1,fparamS2)
   GetFieldProperty=ICORMessage(PadHex(36)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function GetFieldValue(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValue=ICORMessage(PadHex(37)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3))
End Function

Function GetFieldValueByPosition(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueByPosition=ICORMessage(PadHex(38)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3))
End Function

Function GetFieldValueDate(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueDate=str2arr(ICORMessage(PadHex(39)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValueDateTime(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueDateTime=str2arr(ICORMessage(PadHex(40)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValueFloat(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueFloat=CDbl(ICORMessage(PadHex(41)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValueFmt(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueFmt=ICORMessage(PadHex(42)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3))
End Function

Function GetFieldValueInt(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueInt=CLng("&H"+ICORMessage(PadHex(43)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValueLastModification(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueLastModification=str2arr(ICORMessage(PadHex(44)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValuePyTime(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValuePyTime=str2arr(ICORMessage(PadHex(45)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldValueTime(fparamI1,fparamI2,fparamS1,fparamI3)
   GetFieldValueTime=str2arr(ICORMessage(PadHex(46)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetFieldsList(fparamI1,fparamI2)
   GetFieldsList=ICORMessage(PadHex(47)+PadHex(fparamI1)+PadHex(fparamI2))
End Function

Function GetFirstClass(fparamI1)
   GetFirstClass=CLng("&H"+ICORMessage(PadHex(48)+PadHex(fparamI1)))
End Function

Function GetFirstDeletedOffset(fparamI1,fparamI2,fparamS1)
   GetFirstDeletedOffset=CLng("&H"+ICORMessage(PadHex(49)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetFirstFieldValueID(fparamI1,fparamI2,fparamS1)
   GetFirstFieldValueID=CLng("&H"+ICORMessage(PadHex(50)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetFirstObjectID(fparamI1,fparamI2)
   GetFirstObjectID=CLng("&H"+ICORMessage(PadHex(51)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function GetHashFile(fparamS1)
   GetHashFile=ICORMessage(PadHex(52)+PadHex(len(fparamS1))+fparamS1)
End Function

Function GetHashString(fparamS1)
   GetHashString=ICORMessage(PadHex(53)+PadHex(len(fparamS1))+fparamS1)
End Function

Function GetICORHandle()
   GetICORHandle=CLng("&H"+ICORMessage(PadHex(54)))
End Function

Function GetLastFieldValueID(fparamI1,fparamI2,fparamS1)
   GetLastFieldValueID=CLng("&H"+ICORMessage(PadHex(55)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetLastObjectID(fparamI1,fparamI2)
   GetLastObjectID=CLng("&H"+ICORMessage(PadHex(56)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function GetMethodLastModification(fparamI1,fparamI2,fparamS1)
   GetMethodLastModification=str2arr(ICORMessage(PadHex(57)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function GetMethodProperty(fparamI1,fparamI2,fparamS1,fparamS2)
   GetMethodProperty=ICORMessage(PadHex(58)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function GetMethodsList(fparamI1,fparamI2,fparamI3)
   GetMethodsList=ICORMessage(PadHex(59)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3))
End Function

Function GetNextClass(fparamI1,fparamI2)
   GetNextClass=CLng("&H"+ICORMessage(PadHex(60)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function GetNextDeletedOffset(fparamI1,fparamI2,fparamS1,fparamI3)
   GetNextDeletedOffset=CLng("&H"+ICORMessage(PadHex(61)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetNextFieldValueID(fparamI1,fparamI2,fparamS1,fparamI3)
   GetNextFieldValueID=CLng("&H"+ICORMessage(PadHex(62)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetNextFreeObjectID(fparamI1,fparamI2,fparamI3,fparamI4)
   GetNextFreeObjectID=CLng("&H"+ICORMessage(PadHex(130)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(fparamI4)))
End Function

Function GetNextObjectID(fparamI1,fparamI2,fparamI3)
   GetNextObjectID=CLng("&H"+ICORMessage(PadHex(63)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function GetObjectCount(fparamI1,fparamI2)
   GetObjectCount=CLng("&H"+ICORMessage(PadHex(64)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function GetObjectIDByPosition(fparamI1,fparamI2,fparamI3)
   GetObjectIDByPosition=CLng("&H"+ICORMessage(PadHex(65)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function GetObjectModification(fparamI1,fparamI2,fparamI3)
   GetObjectModification=ICORMessage(PadHex(66)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3))
End Function

Function GetPrevFieldValueID(fparamI1,fparamI2,fparamS1,fparamI3)
   GetPrevFieldValueID=CLng("&H"+ICORMessage(PadHex(67)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetPrevObjectID(fparamI1,fparamI2,fparamI3)
   GetPrevObjectID=CLng("&H"+ICORMessage(PadHex(68)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function GetRecLastModification(fparamI1,fparamI2,fparamS1,fparamI3)
   GetRecLastModification=str2arr(ICORMessage(PadHex(69)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetRecOID(fparamI1,fparamI2,fparamS1,fparamI3)
   GetRecOID=CLng("&H"+ICORMessage(PadHex(70)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetRecOwnerID(fparamI1,fparamI2,fparamS1,fparamI3)
   GetRecOwnerID=CLng("&H"+ICORMessage(PadHex(71)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetRecValueAsString(fparamI1,fparamI2,fparamS1,fparamI3)
   GetRecValueAsString=ICORMessage(PadHex(72)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3))
End Function

Function GetStdDialogValue(fparamI1,fparamS1,fparamS2,fparamS3)
   GetStdDialogValue=ICORMessage(PadHex(73)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3)
End Function

Function GetStringAsSafeScriptString(fparamS1)
   GetStringAsSafeScriptString=ICORMessage(PadHex(74)+PadHex(len(fparamS1))+fparamS1)
End Function

Function GetSystemID(fparamS1)
   GetSystemID=ICORMessage(PadHex(75)+PadHex(len(fparamS1))+fparamS1)
End Function

Function GetValueIDByPosition(fparamI1,fparamI2,fparamS1,fparamI3)
   GetValueIDByPosition=CLng("&H"+ICORMessage(PadHex(76)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function GetVariable(fparamI1,fparamS1)
   GetVariable=ICORMessage(PadHex(77)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ICORCompareText(fparamS1,fparamS2)
   ICORCompareText=CLng("&H"+ICORMessage(PadHex(78)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2))
End Function

Function ICORLowerCase(fparamS1)
   ICORLowerCase=ICORMessage(PadHex(79)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ICORSetClipboard(fparamS1)
   ICORMessage(PadHex(80)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ICORUpperCase(fparamS1)
   ICORUpperCase=ICORMessage(PadHex(81)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ImportModuleAsString(fparamI1,fparamS1)
   ImportModuleAsString=ICORMessage(PadHex(82)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1)
End Function

Function IsFieldInClass(fparamI1,fparamI2,fparamS1)
   IsFieldInClass=CLng("&H"+ICORMessage(PadHex(83)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function IsMethodInClass(fparamI1,fparamI2,fparamS1)
   IsMethodInClass=CLng("&H"+ICORMessage(PadHex(84)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function IsMethodInThisClass(fparamI1,fparamI2,fparamS1)
   IsMethodInThisClass=CLng("&H"+ICORMessage(PadHex(85)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1))
End Function

Function IsObjectDeleted(fparamI1,fparamI2,fparamI3)
   IsObjectDeleted=CLng("&H"+ICORMessage(PadHex(86)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function MessageShow(fparamI1,fparamS1,fparamI2,fparamI3)
   MessageShow=CLng("&H"+ICORMessage(PadHex(87)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function ObjectExists(fparamI1,fparamI2,fparamI3)
   ObjectExists=CLng("&H"+ICORMessage(PadHex(88)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)))
End Function

Function OnStdErrPrint(fparamI1,fparamS1,fparamI2)
   ICORMessage(PadHex(89)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI2))
End Function

Function OnStdOutPrint(fparamI1,fparamS1,fparamI2)
   ICORMessage(PadHex(90)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI2))
End Function

Function RepositoryChange(fparamI1,fparamS1,fparamI2,fparamI3,fparamS2,fparamS3,fparamS4)
   ICORMessage(PadHex(91)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3+PadHex(len(fparamS4))+fparamS4)
End Function

Function SelectClassFieldProperties(fparamI1,fparamI2)
   SelectClassFieldProperties=ICORMessage(PadHex(92)+PadHex(fparamI1)+PadHex(fparamI2))
End Function

Function SelectElementDialog(fparamI1,fparamS1,fparamI2)
   SelectElementDialog=ICORMessage(PadHex(93)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI2))
End Function

Function SelectFieldValues(fparamI1,fparamI2,fparamS1)
   SelectFieldValues=ICORMessage(PadHex(94)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1)
End Function

Function SelectInEditor(fparamI1,fparamI2,fparamS1,fparamS2)
   SelectInEditor=ICORMessage(PadHex(95)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function SelectObjects(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4)
   SelectObjects=ICORMessage(PadHex(96)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4))
End Function

Function SelectObjectsFromDictionary(fparamI1,fparamI2,fparamS1,fparamS2)
   SelectObjectsFromDictionary=ICORMessage(PadHex(97)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function SelectSearchInClass(fparamI1,fparamI2)
   SelectSearchInClass=ICORMessage(PadHex(98)+PadHex(fparamI1)+PadHex(fparamI2))
End Function

Function SelectSummaries(fparamI1,fparamI2)
   SelectSummaries=ICORMessage(PadHex(99)+PadHex(fparamI1)+PadHex(fparamI2))
End Function

Function SetClassLastModification(fparamI1,fparamI2,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9)
   ICORMessage(PadHex(100)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9))
End Function

Function SetClassProperty(fparamI1,fparamI2,fparamS1,fparamS2)
   ICORMessage(PadHex(101)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function SetFieldLastModification(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9)
   ICORMessage(PadHex(103)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9))
End Function

Function SetFieldModification(fparamI1,fparamI2,fparamS1,fparamI3,fparamS2)
   ICORMessage(PadHex(104)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(len(fparamS2))+fparamS2)
End Function

Function SetFieldProperty(fparamI1,fparamI2,fparamS1,fparamS2,fparamS3)
   ICORMessage(PadHex(105)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3)
End Function

Function SetFieldValue(fparamI1,fparamI2,fparamS1,fparamI3,fparamS2)
   ICORMessage(PadHex(106)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(len(fparamS2))+fparamS2)
End Function

Function SetFieldValueDate(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6)
   ICORMessage(PadHex(107)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6))
End Function

Function SetFieldValueDateTime(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9,fparamI10)
   ICORMessage(PadHex(108)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9)+PadHex(fparamI10))
End Function

Function SetFieldValueLastModification(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9,fparamI10)
   ICORMessage(PadHex(109)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9)+PadHex(fparamI10))
End Function

Function SetFieldValueTime(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7)
   ICORMessage(PadHex(110)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7))
End Function

Function SetMethodLastModification(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamI5,fparamI6,fparamI7,fparamI8,fparamI9)
   ICORMessage(PadHex(111)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)+PadHex(fparamI6)+PadHex(fparamI7)+PadHex(fparamI8)+PadHex(fparamI9))
End Function

Function SetMethodProperty(fparamI1,fparamI2,fparamS1,fparamS2,fparamS3)
   ICORMessage(PadHex(112)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3)
End Function

Function SetObjectModification(fparamI1,fparamI2,fparamI3,fparamS1)
   ICORMessage(PadHex(113)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(len(fparamS1))+fparamS1)
End Function

Function SetObjectModified(fparamI1,fparamI2,fparamI3)
   ICORMessage(PadHex(114)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3))
End Function

Function SetProgress(fparamI1,fparamI2,fparamI3)
   ICORMessage(PadHex(115)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(fparamI3))
End Function

Function SetTestDecFieldValue(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamS2)
   SetTestDecFieldValue=CLng("&H"+ICORMessage(PadHex(129)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(len(fparamS2))+fparamS2))
End Function

Function SetTestFieldValue(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamS2,fparamS3)
   SetTestFieldValue=CLng("&H"+ICORMessage(PadHex(127)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3))
End Function

Function SetTestIncFieldValue(fparamI1,fparamI2,fparamS1,fparamI3,fparamI4,fparamS2)
   SetTestIncFieldValue=CLng("&H"+ICORMessage(PadHex(128)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(len(fparamS2))+fparamS2))
End Function

Function SetVariable(fparamI1,fparamS1,fparamS2)
   ICORMessage(PadHex(116)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2)
End Function

Function ShellExecute(fparamI1,fparamS1,fparamS2,fparamS3)
   ShellExecute=CLng("&H"+ICORMessage(PadHex(117)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3))
End Function

Function SortFile(fparamI1,fparamS1,fparamS2,fparamI2,fparamI3,fparamI4,fparamI5)
   SortFile=CLng("&H"+ICORMessage(PadHex(118)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(fparamI2)+PadHex(fparamI3)+PadHex(fparamI4)+PadHex(fparamI5)))
End Function

Function StatusInfo(fparamI1,fparamS1)
   ICORMessage(PadHex(119)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1)
End Function

Function Str2DateTime(fparamS1)
   Str2DateTime=str2arr(ICORMessage(PadHex(120)+PadHex(len(fparamS1))+fparamS1))
End Function

Function Str2HTMLStr(fparamS1)
   Str2HTMLStr=ICORMessage(PadHex(121)+PadHex(len(fparamS1))+fparamS1)
End Function

Function SummaryEdit(fparamI1,fparamI2)
   SummaryEdit=CLng("&H"+ICORMessage(PadHex(122)+PadHex(fparamI1)+PadHex(fparamI2)))
End Function

Function URLString2NormalString(fparamS1)
   URLString2NormalString=ICORMessage(PadHex(123)+PadHex(len(fparamS1))+fparamS1)
End Function

Function ValueExists(fparamI1,fparamI2,fparamS1,fparamI3)
   ValueExists=CLng("&H"+ICORMessage(PadHex(124)+PadHex(fparamI1)+PadHex(fparamI2)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI3)))
End Function

Function getcommand(fparamI1,fparamS1,fparamS2,fparamS3)
   getcommand=ICORMessage(PadHex(125)+PadHex(fparamI1)+PadHex(len(fparamS1))+fparamS1+PadHex(len(fparamS2))+fparamS2+PadHex(len(fparamS3))+fparamS3)
End Function

Function getline(fparamS1,fparamI1)
   getline=ICORMessage(PadHex(126)+PadHex(len(fparamS1))+fparamS1+PadHex(fparamI1))
End Function
%>
