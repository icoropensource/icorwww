<HTML>
<HEAD>
<STYLE TYPE="text/css">
<!--
  A:link { color: "#000080"; }
  A:visited { color: "#000000"; }
BODY  { background: teal;
   color: white;
   font-size: 10pt;
   font-family: Verdana, Arial, Helvetica, sans-serif }
P  { color: black;
   font-size: 10pt;
   font-family: Verdana, Arial, Helvetica, sans-serif } 
P.redintro  { color: #ff3300;
   font-size: 11pt;
   line-height: 15pt;
   font-weight: bold;
   font-family: Verdana, Arial, Helvetica, sans-serif } 
P.blueintro { color: #0099ff;
   font-size: 11pt;
   line-height: 15pt;
   font-weight: bold;
   font-family: Verdana, Arial, Helvetica, sans-serif } 
P.note   { color: black;
   font-size: 8pt;
   font-family: Verdana, Arial, Helvetica, sans-serif } 
.indent  { color: black;
   font-size: 10pt;
   text-indent: 15pt;
   font-family: Verdana, Arial, Helvetica, sans-serif } 
H1, H2 { font-size: 18pt; 
   font-weight: bold;
   color: white; 
   font-family: Verdana, Arial, Helvetica, sans-serif }
H3, H4 { font-size: 12pt; 
   font-weight: bold;
   color: white; 
   font-family: Verdana, Arial, Helvetica, sans-serif }
-->
</STYLE>
<TITLE>Debug</TITLE>
</HEAD>
<BODY>
<BLOCKQUOTE>

<H1>Debug Information</H1>

<H3>Application Contents</H3>

<%
On Error Resume Next

  For Each Key In Application.Contents
    Response.Write Key & " = "
    If IsObject(Application.Contents(Key)) Then
      Response.Write "<i>(object)</i>" & "<BR>"
    Else
      Response.Write CStr(Application.Contents(Key)) & "<BR>"
    End If
  Next
%>

<HR>

<H3>Application Static Objects</H3>
<%
  For Each Key In Application.StaticObjects
    Response.Write Key & " = <i>(object)</i><BR>"
  Next
%>

<HR>

<H3>Request Client Certificate</H3>
<%
  For Each Key In Request.ClientCertificate
    Response.Write Key & " = " & Request.ClientCertificate(Key) & "<BR>"
  Next
%>

<HR>

<H3>Request Cookies</H3>

<%
  For Each Cookie In Request.Cookies
    If Request.Cookies(Cookie).HasKeys Then
      For Each Key In Request.Cookies(Cookie)
        Response.Write Cookie & "(" & Key & ") = " & _
          Request.Cookies(Cookie)(Key) & "<BR>"
      Next
    Else
      Response.Write Cookie & " = " & Request.Cookies(Cookie) & "<BR>"
    End If
  Next
%>

<HR>

<H3>Request Form</H3>

<%
  For Each Key In Request.Form
    Response.Write Key & " = " & Request.Form(Key) & "<BR>"
  Next
%>

<HR>

<H3>Request Query String</H3>

<%
  For Each Key In Request.QueryString
    Response.Write Key & " = " & Request.QueryString(Key) & "<BR>"
  Next
%>

<HR>

<H3>Request Server Variables</H3>

<%
  For Each Key In Request.ServerVariables
    Response.Write Key & " = " & Request.ServerVariables(Key) & "<BR>"
  Next
%>

<HR>

<H3>Session Static Objects</H3>

<%
  For Each Key In Session.StaticObjects
    Response.Write Key & " = <i>(object)</i><BR>"
  Next
%>

<HR>

<H3>Session Contents</H3>

<%
  For Each Key In Session.Contents
    Response.Write Key & " = " & Session.Contents(Key) & "<BR>"
    If IsObject(Session.Contents(Key)) Then
      Response.Write "<i>(object)</i>" & "<BR>"
    Else
      Response.Write Session.Contents(Key) & "<BR>"
    End If
  Next
%>

</BLOCKQUOTE>
</BODY>
</HTML>