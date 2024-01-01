<%
 ' FCKeditor - The text editor for Internet - http://www.fckeditor.net
 ' Copyright (C) 2003-2008 Frederico Caldeira Knabben
 '
 ' == BEGIN LICENSE ==
 '
 ' Licensed under the terms of any of the following licenses at your
 ' choice:
 '
 '  - GNU General Public License Version 2 or later (the "GPL")
 '    http://www.gnu.org/licenses/gpl.html
 '
 '  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
 '    http://www.gnu.org/licenses/lgpl.html
 '
 '  - Mozilla Public License Version 1.1 or later (the "MPL")
 '    http://www.mozilla.org/MPL/MPL-1.1.html
 '
 ' == END LICENSE ==
 '
 ' This file include IO specific functions used by the ASP Connector.
%>
<%
function CombinePaths( sBasePath, sFolder)
	CombinePaths =  RemoveFromEnd( sBasePath, "/" ) & "/" & RemoveFromStart( sFolder, "/" )
end function

function CombineLocalPaths( sBasePath, sFolder)
	sFolder = replace(sFolder, "/", "\")
	' The RemoveFrom* functions use RegExp, so we must escape the \
	CombineLocalPaths =  RemoveFromEnd( sBasePath, "\\" ) & "\" & RemoveFromStart( sFolder, "\\" )
end function

Function GetResourceTypePath( resourceType, sCommand )
	if ( sCommand = "QuickUpload") then
		GetResourceTypePath = ConfigQuickUploadPath.Item( resourceType )
	else
		GetResourceTypePath = ConfigFileTypesPath.Item( resourceType )
	end if
end Function

Function GetResourceTypeDirectory( resourceType, sCommand )
	if ( sCommand = "QuickUpload") then

		if ( ConfigQuickUploadAbsolutePath.Item( resourceType ) <> "" ) then
			GetResourceTypeDirectory = ConfigQuickUploadAbsolutePath.Item( resourceType )
		else
			' Map the "UserFiles" path to a local directory.
			GetResourceTypeDirectory = Server.MapPath( ConfigQuickUploadPath.Item( resourceType ) )
		end if
	else
		if ( ConfigFileTypesAbsolutePath.Item( resourceType ) <> "" ) then
			GetResourceTypeDirectory = ConfigFileTypesAbsolutePath.Item( resourceType )
		else
			' Map the "UserFiles" path to a local directory.
			GetResourceTypeDirectory = Server.MapPath( ConfigFileTypesPath.Item( resourceType ) )
		end if
	end if
end Function

Function GetUrlFromPath( resourceType, folderPath, sCommand )
	GetUrlFromPath = CombinePaths( GetResourceTypePath( resourceType, sCommand ), folderPath )
End Function

Function RemoveExtension( fileName )
	RemoveExtension = Left( fileName, InStrRev( fileName, "." ) - 1 )
End Function

Function ServerMapFolder( resourceType, folderPath, sCommand )
	Dim sResourceTypePath
	' Get the resource type directory.
	sResourceTypePath = GetResourceTypeDirectory( resourceType, sCommand )

	' Ensure that the directory exists.
	CreateServerFolder sResourceTypePath

	' Return the resource type directory combined with the required path.
	ServerMapFolder = CombineLocalPaths( sResourceTypePath, folderPath )
End Function

Sub CreateServerFolder( folderPath )
	Dim oFSO
	Set oFSO = Server.CreateObject( "Scripting.FileSystemObject" )

	Dim sParent
	sParent = oFSO.GetParentFolderName( folderPath )

	' If folderPath is a network path (\\server\folder\) then sParent is an empty string.
	' Get out.
	if (sParent = "") then exit sub

	' Check if the parent exists, or create it.
	If ( NOT oFSO.FolderExists( sParent ) ) Then CreateServerFolder( sParent )

	If ( oFSO.FolderExists( folderPath ) = False ) Then
		On Error resume next
		oFSO.CreateFolder( folderPath )

		if err.number<>0 then
		dim sErrorNumber
		Dim iErrNumber, sErrDescription
		iErrNumber		= err.number
		sErrDescription	= err.Description

		On Error Goto 0

		Select Case iErrNumber
			Case 52
				sErrorNumber = "102"	' Invalid Folder Name.
			Case 70
				sErrorNumber = "103"	' Security Error.
			Case 76
				sErrorNumber = "102"	' Path too long.
			Case Else
				sErrorNumber = "110"
			End Select

			SendError sErrorNumber, "CreateServerFolder(" & folderPath & ") : " & sErrDescription
		end if

	End If

	Set oFSO = Nothing
End Sub

Function IsAllowedExt( extension, resourceType )
	Dim oRE
	Set oRE	= New RegExp
	oRE.IgnoreCase	= True
	oRE.Global		= True

	Dim sAllowed, sDenied
	sAllowed	= ConfigAllowedExtensions.Item( resourceType )
	sDenied		= ConfigDeniedExtensions.Item( resourceType )

	IsAllowedExt = True

	If sDenied <> "" Then
		oRE.Pattern	= sDenied
		IsAllowedExt	= Not oRE.Test( extension )
	End If

	If IsAllowedExt And sAllowed <> "" Then
		oRE.Pattern		= sAllowed
		IsAllowedExt	= oRE.Test( extension )
	End If

	Set oRE	= Nothing
End Function

Function IsAllowedType( resourceType )
	Dim oRE
	Set oRE	= New RegExp
	oRE.IgnoreCase	= True
	oRE.Global		= True
	oRE.Pattern		= "^(" & ConfigAllowedTypes & ")$"

	IsAllowedType = oRE.Test( resourceType )

	Set oRE	= Nothing
End Function

Function IsAllowedCommand( sCommand )
	Dim oRE
	Set oRE	= New RegExp
	oRE.IgnoreCase	= True
	oRE.Global		= True
	oRE.Pattern		= "^(" & ConfigAllowedCommands & ")$"

	IsAllowedCommand = oRE.Test( sCommand )

	Set oRE	= Nothing
End Function

function GetCurrentFolder()
	dim sCurrentFolder
	sCurrentFolder = Request.QueryString("CurrentFolder")
	If ( sCurrentFolder = "" ) Then sCurrentFolder = "/"

	' Check the current folder syntax (must begin and start with a slash).
	If ( Right( sCurrentFolder, 1 ) <> "/" ) Then sCurrentFolder = sCurrentFolder & "/"
	If ( Left( sCurrentFolder, 1 ) <> "/" ) Then sCurrentFolder = "/" & sCurrentFolder

	' Check for invalid folder paths (..)
	If ( InStr( 1, sCurrentFolder, ".." ) <> 0 OR InStr( 1, sCurrentFolder, "\" ) <> 0) Then
		SendError 102, ""
	End If

	GetCurrentFolder = sCurrentFolder
end function

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
   s=RegExReplace(s,"(ą)", "a")
   s=RegExReplace(s,"(ć)", "c")
   s=RegExReplace(s,"(ę)", "e")
   s=RegExReplace(s,"(ł)", "l")
   s=RegExReplace(s,"(ń)", "n")
   s=RegExReplace(s,"(ó)", "o")
   s=RegExReplace(s,"(ś)", "s")
   s=RegExReplace(s,"(ż|ź)", "z")   
   s=RegExReplace(s,"(Ą)", "A")
   s=RegExReplace(s,"(Ć)", "C")
   s=RegExReplace(s,"(Ę)", "E")
   s=RegExReplace(s,"(Ł)", "L")
   s=RegExReplace(s,"(Ń)", "N")
   s=RegExReplace(s,"(Ó)", "O")
   s=RegExReplace(s,"(Ś)", "S")
   s=RegExReplace(s,"(Ż|Ź)", "Z")   
   
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

' Do a cleanup of the folder name to avoid possible problems
function SanitizeFolderName( sNewFolderName )
	Dim oRegex
	Set oRegex = New RegExp
	oRegex.Global = True
	oRegex.Pattern = "( |\.|\\|\/|\||:|\?|\*|""|\<|\>|[\u0000-\u001F]|\u007F)"
	SanitizeFolderName = oRegex.Replace( sNewFolderName, "_" )
   SanitizeFolderName=UTF2NoPL(SanitizeFolderName)
	Set oRegex = Nothing
end function

' Do a cleanup of the file name to avoid possible problems
function SanitizeFileName( sNewFileName )
	Dim oRegex
	Set oRegex = New RegExp
	oRegex.Global = True

	if ( ConfigForceSingleExtension = True ) then
		oRegex.Pattern = "\.(?![^.]*$)"
		sNewFileName = oRegex.Replace( sNewFileName, "_" )
	end if
	oRegex.Pattern = "( |\\|\/|\||:|\?|\*|""|\<|\>|[\u0000-\u001F]|\u007F)"
	SanitizeFileName = oRegex.Replace( sNewFileName, "_" )
   SanitizeFileName=UTF2NoPL(SanitizeFileName)
	Set oRegex = Nothing
end function

' This is the function that sends the results of the uploading process.
Sub SendUploadResults( errorNumber, fileUrl, fileName, customMsg )
   dim aMsgOut, ufn
	Response.Clear
	Response.Write "<script type=""text/javascript"">"


   Response.Write "function htmlDecode(str) {"
   Response.Write "  var entMap={'quot':34,'amp':38,'apos':39,'lt':60,'gt':62};"
   Response.Write "  return str.replace(/&([^;]+);/g,function(m,n) {"
   Response.Write "    var code;"
   Response.Write "    if (n.substr(0,1)=='#') {"
   Response.Write "      if (n.substr(1,1)=='x') {"
   Response.Write "        code=parseInt(n.substr(2),16);"
   Response.Write "      } else {"
   Response.Write "        code=parseInt(n.substr(1),10);"
   Response.Write "      }"
   Response.Write "    } else {"
   Response.Write "      code=entMap[n];"
   Response.Write "    }"
   Response.Write "    return (code===undefined||code===NaN)?'&'+n+';':String.fromCharCode(code);"
   Response.Write "  });"
   Response.Write "}"


   'Response.Write  "prompt('','" & fileUrl & "');"
	' Minified version of the document.domain automatic fix script (#1919).
	' The original script can be found at _dev/domain_fix_template.js
   
   Response.Write "if (window.parent.CKEDITOR==undefined) {"
   Response.Write "(function(){var d=document.domain;while (true){try{var A=window.parent.document.domain;break;}catch(e) {};d=d.replace(/.*?(?:\.|$)/,'');if (d.length==0) break;try{document.domain=d;}catch (e){break;}}})();"
   Response.Write "  window.parent.OnUploadCompleted(" & errorNumber & ",""" & Replace( fileUrl, """", "\""" ) & """,""" & Replace( fileName, """", "\""" ) & """,""" & Replace( customMsg , """", "\""" ) & """) ;"
   Response.Write "} else {"
   
   if errorNumber=0 then
      aMsgOut=""
   elseif errorNumber=1 then
      aMsgOut=Replace( customMsg , """", "\""" )
   elseif errorNumber=201 then
      aMsgOut="Plik z taka nazwa juz istnieje. Nazwa zostala zmieniona na: " & fileName
   elseif  errorNumber=202 then
      aMsgOut="Niepoprawny plik"
   else
      aMsgOut="Blad podczas ladowania pliku. Kod bledu: " & errorNumber
   end if      
   
'   Response.Write "window.parent.CKEDITOR.tools.callFunction(1, '" & fileUrl &  "','" & aMsgOut & "');"
'   Response.Write "if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)){"
'   Response.Write "  var uf = 2;"
'   Response.Write "}else{"
'   Response.Write "  var uf = 1;"
'   Response.Write "};"
'   Response.Write "alert(htmlDecode(fileUrl));"
   ufn=request.QueryString("CKEditorFuncNum")
   if ufn="" then
      ufn="1"
   end if

   Response.Write "window.parent.CKEDITOR.tools.callFunction(" & ufn & ", htmlDecode('" & fileUrl &  "'),'" & aMsgOut & "');"
   Response.Write "}"
   Response.Write "</script>"
End Sub

%>
