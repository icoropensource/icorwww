<%
'<!--
' * FCKeditor - The text editor for Internet - http://www.fckeditor.net
' * Copyright (C) 2003-2007 Frederico Caldeira Knabben
' *
' * == BEGIN LICENSE ==
' *
' * Licensed under the terms of any of the following licenses at your
' * choice:
' *
' *  - GNU General Public License Version 2 or later (the "GPL")
' *    http://www.gnu.org/licenses/gpl.html
' *
' *  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
' *    http://www.gnu.org/licenses/lgpl.html
' *
' *  - Mozilla Public License Version 1.1 or later (the "MPL")
' *    http://www.mozilla.org/MPL/MPL-1.1.html
' *
' * == END LICENSE ==
' *
' * These are the classes used to handle ASP upload without using third
' * part components (OCX/DLL).
'-->
'**********************************************
' File:     NetRube_Upload.asp
' Version:  NetRube Upload Class Version 2.3 Build 20070528
' Author:   NetRube
' Email: NetRube@126.com
' Date:     05/28/2007
' Comments: The code for the Upload.
'        This can free usage, but please
'        not to delete this copyright information.
'        If you have a modification version,
'        Please send out a duplicate to me.
'**********************************************

function UTF2Win(s)
   UTF2Win=s
end function

' Allowed and Denied extensions configurations.
Dim ConfigAllowedExtensions, ConfigDeniedExtensions
Set ConfigAllowedExtensions   = CreateObject( "Scripting.Dictionary" )
Set ConfigDeniedExtensions = CreateObject( "Scripting.Dictionary" )

ConfigAllowedExtensions.Add   "File", ""
ConfigDeniedExtensions.Add "File", "html|htm|php|php2|php3|php4|php5|phtml|pwml|inc|asp|aspx|ascx|jsp|cfm|cfc|pl|bat|exe|com|dll|vbs|js|reg|cgi|htaccess|asis|sh|shtml|shtm|phtm"

ConfigAllowedExtensions.Add   "Image", "jpg|gif|jpeg|png|bmp"
ConfigDeniedExtensions.Add "Image", ""

ConfigAllowedExtensions.Add   "Flash", "swf|fla"
ConfigDeniedExtensions.Add "Flash", ""

ConfigAllowedExtensions.Add   "Document", "jpg|gif|jpeg|tif|doc|pdf|docx|odf|rtf|wri|txt"
ConfigDeniedExtensions.Add "Document", ""

Function RemoveExtension( fileName )
   RemoveExtension = Left( fileName, InStrRev( fileName, "." ) - 1 )
End Function

Class NetRube_Upload

   Public   File, Form
   Private oSourceData
   Private nMaxSize, nErr, sAllowed, sDenied, nMultipleFiles
   Private nAutoUTF2Win
   
   Private Sub Class_Initialize
      nErr     = 0

      nAutoUTF2Win=1
      nMultipleFiles = 0
      

      nMaxSize = 150000000 '

      
      Set File       = Server.CreateObject("Scripting.Dictionary")
      File.CompareMode  = 1
      Set Form       = Server.CreateObject("Scripting.Dictionary")
      Form.CompareMode  = 1
      
      Set oSourceData      = Server.CreateObject("ADODB.Stream")
      oSourceData.Type  = 1
      oSourceData.Mode  = 3
      oSourceData.Open
   End Sub
   
   Private Sub Class_Terminate
      Form.RemoveAll
      Set Form = Nothing
      File.RemoveAll
      Set File = Nothing
      
      oSourceData.Close
      Set oSourceData = Nothing
   End Sub
   
   Public Property Get Version
      Version = "NetRube Upload Class Version 2.3 Build 20070528"
   End Property

   Public Property Get ErrNum
      ErrNum   = nErr
   End Property

   Public Property Let AutoUTF2Win(v)
      nAutoUTF2Win = v
   End Property

   Public Property Let MaxSize(nSize)
      nMaxSize = nSize
   End Property

   Public Property Let MultipleFiles(v)
      nMultipleFiles = v
   End Property
   
   Public Property Let Allowed(sExt)
      sAllowed = sExt
   End Property

   Public Property Let Denied(sExt)
      sDenied  = sExt
   End Property

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

   Public Sub GetData
      Dim nTotalSize
      nTotalSize  = Request.TotalBytes
      If nTotalSize < 1 Then
'         nErr = 2
         Exit Sub
      End If
      If nMaxSize > 0 And nTotalSize > nMaxSize Then
         nErr = 3
         Exit Sub
      End If

      Dim aCType
      aCType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")

      If aCType(0) = "application/x-www-form-urlencoded" Then
         Dim SourceData
         SourceData = Request.BinaryRead(Request.TotalBytes)
         SourceData = RSBinaryToString(SourceData)
         SourceData = Split(SourceData, "&")
         Dim Field, FieldName, FieldContents
         For Each Field In SourceData
           Field = Split(Field, "=")
           FieldName = DecodeUrlData(Field(0))
           FieldContents = DecodeUrlData(Field(1))
           
           'if nAutoUTF2Win=1 then
           '   Form.Add UTF2Win(FieldName), UTF2Win(FieldContents)
           'else
           '   Form.Add FieldName, FieldContents
           'end if
           
           '$$SK 20091115 fields with multiple values
           if nAutoUTF2Win=1 then
              FieldName=UTF2Win(FieldName)
              FieldContents=UTF2Win(FieldContents)
           end if
           if Form.Exists(FieldName) then
              Form(FieldName)=Form(FieldName) & " " & FieldContents
           else
              Form.Add FieldName, FieldContents
           end if
         Next
      end if
      If aCType(0) = "multipart/form-data" Then
         'Thankful long(yrl031715@163.com)
         'Fix upload large file.
   
         Dim nTotalBytes, nPartBytes, ReadBytes
         ReadBytes = 0
         nTotalBytes = Request.TotalBytes
         '&#229;&#190;&#170;ÁéØ&#229;_+&#229;&#157;-&#232;Øª&#229;è-
         Do While ReadBytes < nTotalBytes
            '&#229;_+&#229;&#157;-&#232;Øª&#229;è-
            nPartBytes = 64 * 1024 '&#229;_+Ê__ÊØè&#229;&#157;-64k
            If nPartBytes + ReadBytes > nTotalBytes Then 
               nPartBytes = nTotalBytes - ReadBytes
            End If
            oSourceData.Write Request.BinaryRead(nPartBytes)
            ReadBytes = ReadBytes + nPartBytes
         Loop
         '**********************************************
         oSourceData.Position = 0
         
         Dim oTotalData, oFormStream, sFormHeader, sFormName, bCrLf, nBoundLen, nFormStart, nFormEnd, nPosStart, nPosEnd, sBoundary
         
         oTotalData  = oSourceData.Read
         bCrLf    = ChrB(13) & ChrB(10)
         sBoundary   = MidB(oTotalData, 1, InStrB(1, oTotalData, bCrLf) - 1)
         nBoundLen   = LenB(sBoundary) + 2
         nFormStart  = nBoundLen
         
         Set oFormStream = Server.CreateObject("ADODB.Stream")
         
         Do While (nFormStart + 2) < nTotalSize
            nFormEnd = InStrB(nFormStart, oTotalData, bCrLf & bCrLf) + 3
            
            With oFormStream
               .Type = 1
               .Mode = 3
               .Open
               oSourceData.Position = nFormStart
               oSourceData.CopyTo oFormStream, nFormEnd - nFormStart
               .Position   = 0
               .Type    = 2
               '.CharSet = "utf-8"
               .CharSet = "utf-8"
               sFormHeader = .ReadText
               .Close
            End With
            
            nFormStart  = InStrB(nFormEnd, oTotalData, sBoundary) - 1
            nPosStart   = InStr(22, sFormHeader, " name=", 1) + 7
            nPosEnd     = InStr(nPosStart, sFormHeader, """")
            sFormName   = Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
            
            If InStr(45, sFormHeader, " filename=", 1) > 0 Then
               If nMultipleFiles Then
                  Dim nfi,bFormName
                  For nfi=1 To 100
                     bFormName=sFormName+"_"+CStr(nfi)
                     If Not File.Exists(bFormName) Then
                        sFormName=bFormName
                        Exit For
                     End If
                  Next
               End If
               Set File(sFormName)        = New NetRube_FileInfo
               File(sFormName).FormName   = sFormName
               File(sFormName).Start      = nFormEnd
               File(sFormName).Size    = nFormStart - nFormEnd - 2
               nPosStart               = InStr(nPosEnd, sFormHeader, " filename=", 1) + 11
               nPosEnd                 = InStr(nPosStart, sFormHeader, """")
               File(sFormName).ClientPath = Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
               File(sFormName).Name    = Mid(File(sFormName).ClientPath, InStrRev(File(sFormName).ClientPath, "\") + 1)
               File(sFormName).Ext        = LCase(Mid(File(sFormName).Name, InStrRev(File(sFormName).Name, ".") + 1))
               nPosStart               = InStr(nPosEnd, sFormHeader, "Content-Type: ", 1) + 14
               nPosEnd                 = InStr(nPosStart, sFormHeader, vbCr)
               File(sFormName).MIME    = Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
            Else
               With oFormStream
                  .Type = 1
                  .Mode = 3
                  .Open
                  oSourceData.Position = nFormEnd
                  oSourceData.CopyTo oFormStream, nFormStart - nFormEnd - 2
                  .Position   = 0
                  .Type    = 2
'                  .CharSet = "utf-8"
                  .CharSet = "utf-8"
'                  Form(UTF2Win(sFormName))   = UTF2Win(.ReadText)
                  Form(sFormName)   = .ReadText
'                  Response.Write "FF: " + sFormName + " ["+Form(sFormName)+"]{"+UTF2Win(Form(sFormName))+"}"+chr(13)+chr(10)
                  .Close
               End With
            End If
            
            nFormStart  = nFormStart + nBoundLen
         Loop
         
         oTotalData = ""
         Set oFormStream = Nothing

      End If
      
   End Sub

   Public Sub SaveAs(sItem, sFileName)
      If File(sItem).Size < 1 Then
         nErr = 2
         Exit Sub
      End If
      
      If Not IsAllowed(File(sItem).Ext) Then
         nErr = 4
         Exit Sub
      End If
      
      Dim oFileStream
      Set oFileStream = Server.CreateObject("ADODB.Stream")
      With oFileStream
         .Type    = 1
         .Mode    = 3
         .Open
         oSourceData.Position = File(sItem).Start
         oSourceData.CopyTo oFileStream, File(sItem).Size
         .Position   = 0
         .SaveToFile sFileName, 2
         .Close
      End With
      Set oFileStream = Nothing
   End Sub
   
   Private Function IsAllowed(sExt)
      Dim oRE
      Set oRE  = New RegExp
      oRE.IgnoreCase = True
      oRE.Global     = True
      
      If sDenied = "" Then
         oRE.Pattern = sAllowed
         IsAllowed   = (sAllowed = "") Or oRE.Test(sExt)
      Else
         oRE.Pattern = sDenied
         IsAllowed   = Not oRE.Test(sExt)
      End If
      
      Set oRE  = Nothing
   End Function
End Class

Class NetRube_FileInfo
   Dim FormName, ClientPath, Path, Name, Ext, Content, Size, MIME, Start
End Class
%>