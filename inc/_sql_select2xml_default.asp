<!-- #include file="adovbs.inc" -->
<%
aservertype="IIS"
aserverversion="5"

if aservertype<>"IIS" then
   Response.Write "<html><body><h1>Wymagany serwer IIS w wersji 4 lub wyzszej.</h1></body></html>"
   Response.End
end if

Response.ContentType = "text/xml"

rs.ActiveConnection=cn
rs.Source=aselect
rs.Open

if not rs.EOF then
   Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>"
   if (aserverversion="4") or (aserverversion="5") then
      Set stm=Server.CreateObject("ADODB.Stream")
      stm.Type=adTypeText
      '  stm.Charset="utf-8"
      rs.Save stm, adPersistXML
      Response.Write stm.ReadText
      stm.Close
      Set stm = Nothing
   end if
   if aserverversion="5?" then
      rs.Save Response, adPersistXML
   end if
else
   Response.Write "<root>Brak danych</root>"
end if

if rs.State<>adStateClosed then
   rs.Close
end if
Set rs=Nothing
Set cn=Nothing
%>
