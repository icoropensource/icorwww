<%@EnableSessionState=False%>
<!--#INCLUDE FILE="_upload.asp"-->
<%
'Session must be off to work correctly.

Const RefreshTime = 1'Seconds

'Upload ID must be defined.
'Redirect to base script without the parameter.
'if Request.QueryString("UploadID") = "" then 
'  response.redirect "Upload-Progress.ASP"
'end if
Server.ScriptTimeout = 10

'Simple utilities.
'Formats milisedond to m:ss format.
Function FormatTime(byval ms)
  ms = 0.001 * ms 'get second
  FormatTime = (ms \ 60) & ":" & right("0" & (ms mod 60),2) & "s"
End Function 

'Format bytes to a string
Function FormatSize(byval Number)
  if isnumeric(Number) then
    if Number > &H100000 then'1M
      Number = FormatNumber (Number/&H100000,1) & "MB"
    elseif Number > 1024 then'1k
      Number = FormatNumber (Number/1024,1) & "kB"
    else
      Number = FormatNumber (Number,0) & "B"
    end if
  end if
  FormatSize = Number
End Function

Function FileStateInfo(Form)
  'enumerate uploaded fields.
  'and build report about its current state.
   On Error Resume Next
  Dim UpStateHTML, Field
  for each Field in Form.Files
    'Get field name
    UpStateHTML = UpStateHTML & "Plik:" & Field.Name
    if Field.InProgress then
      'this field is in progress now.
      UpStateHTML = UpStateHTML & ", transfer: " & Field.FileName
    elseif Field.Length>0 then
      'This field was succefully uploaded.
      UpStateHTML = UpStateHTML & ", zapisany: " & Field.FileName & ", " & FormatSize(Field.Length)
    end if
    UpStateHTML = UpStateHTML & "<br>"
  Next
  FileStateInfo = UpStateHTML
End Function

Dim Form
Set Form = New ASPForm 
  '{b}Get current uploading form with UploadID.
on error resume next
Set Form = Form.getForm(Request.QueryString("UploadID"))'{/b}
  
if Err = 0 then '?Completted 0 = in progress
  on error goto 0
  if Form.BytesRead>0 then'Upload was started.
    Dim UpStateHTML
    UpStateHTML = FileStateInfo(Form) 'Get currently uploadded filenames and sizes
  end if

  response.cachecontrol = "no-cache"
  response.AddHeader "Pragma","no-cache"
  response.addheader "Refresh", RefreshTime

  'Count progress indicators
  ' - percent and total read, total bytes, etc.
  Dim PercBytesRead, PercentRest, BytesRead, TotalBytes
  Dim UploadTime, RestTime, TransferRate, InServerSaving
  BytesRead = Form.BytesRead
  TotalBytes = Form.TotalBytes
  if (BytesRead>=TotalBytes) and (TotalBytes>0) then
    InServerSaving=1
  else
    InServerSaving=0
  end if

  if TotalBytes>0 then 
    'upload started.
    PercBytesRead = int(100*BytesRead/TotalBytes)
    PercentRest = 100-PercBytesRead
    
    if Form.ReadTime>0 then 
      TransferRate = BytesRead / Form.ReadTime
    end if
    if TransferRate>0 then 
      RestTime = FormatTime((TotalBytes-BytesRead) / TransferRate) 
    end if
    TransferRate = FormatSize(1000 * TransferRate)
  else
    'upload not started.
    RestTime = "?"
    PercBytesRead = 0
    PercentRest = 100
    TransferRate = "?"
  end if

  'Create graphics progress bar.
  'The bar is created with blue (TDsread, completted) / blank (TDsRemain, remaining) TD cells.
  Dim TDsread, TDsRemain
  TDsread = replace(space(0.5*PercBytesRead), " ", "<TD BGColor=blue Class=p> </TD>")
  TDsRemain = replace(space(0.5*PercentRest), " ", "<TD Class=p> </TD>")

  'Format output values.
  UploadTime = FormatTime(Form.ReadTime)
  TotalBytes = FormatSize(TotalBytes)
  BytesRead = FormatSize(BytesRead)
%>
<HTML>
<Head>
 <style type='text/css'>
BODY
{
    BACKGROUND: ivory;
    COLOR: black;
    FONT: 8pt Arial;
    MARGIN: 0.1in;
}
  TD{font-size:9pt} 
  TD.p{font-size:6px;Height:6px;border:1px inset white}
 </style>
 <meta http-equiv="Page-Enter" content="revealTrans(Duration=0,Transition=6)"> 
 <META HTTP-EQUIV="Refresh" CONTENT="<%=RefreshTime%>">
 <Title><%=PercBytesRead%>% ukoñczono</Title>
</Head>

<Body BGcolor=Silver LeftMargin=15 TopMargin=4 RIGHTMARGIN=4 BOTTOMMARGIN=4>

<%
   if InServerSaving=0 then
%>
<Table cellpadding=0 cellspacing=0 border=0 width=100% >
<tr>
 <%=TDsread%><%=TDsRemain%>
</tr>
</table>

<Center>
<br>
<Table cellpadding=0 cellspacing=5 border=0>
<tr>
 <Td align="right"><b>Postêp:</b></td>
 <Td><%=BytesRead%> z <%=TotalBytes%> (<%=PercBytesRead%>%) </Td>
</tr>
<tr>
 <Td align="right"><b>Czas:</b></td>
 <Td><%=UploadTime%> (<%=TransferRate%>/s) </Td>
</tr>
<tr>
 <Td align="right"><b>Do zakoñczenia:</b></td>
 <Td><%=RestTime%> </Td>
</tr>
</table>
<br>
<Input Type=Button Value="Anuluj" OnClick="Cancel()" Style="background-color:red;color:white;cursor:hand;font-weight:bold">
</Center>
<br>
<%
   Else
%>

<Table cellpadding=0 cellspacing=0 border=0 width=100% >
<tr>
<td><center><br><br><font color="navy" size="+2"><b>Proszê czekaæ, trwa zapisywanie na serwerze.</b></font></center></td>
</tr>
</Table>

<%
   End If
%>

<%=UpStateHTML%>

<Script>
//I'm sorry. IE enables switch-off refresh header. You can use script to do the same action
//to be sure that progress window will refresh.
window.setTimeout('refresh()', <%=RefreshTime%>*2000);

function refresh(){
   window.location.href = window.location.href;
   window.setTimeout('refresh()', <%=RefreshTime%>*2000);
}
function Cancel(){
   //get opener location - this is URL of the main upload script.
   var l = ''+opener.document.location;
   
   //Add Action=Cancel querystring parameter
   if (l.indexOf('Action=Cancel')<0) {
      l += (l.indexOf('?')<0 ? '?' : '&') + 'Action=Cancel'
   };

   //Set the new URL to opener (upload is cancelled)  
   opener.document.location = l;

   //Close this window.
   window.close();
}
</Script>

</Body>
</HTML>

<%
else 'if Err = 0 then upload finished
%>
<HTML>
 <HEAD>
 <TITLE>Zapis ukoñczony</TITLE>
 <Script>window.close();</Script>
 </HEAD>
</HTML>
<%
End If'if Err = 0 then 
%>
