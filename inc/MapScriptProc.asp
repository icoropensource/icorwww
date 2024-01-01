<%@ CodePage=65001 LANGUAGE="VBSCRIPT" %>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:t="urn:schemas-microsoft-com:time" xmlns="http://www.w3.org/TR/REC-html40" xmlns:tool>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="pl">
<meta name="pragma" content="no-cache">
<meta name="keyword" content="¹æê³ñóœŸ¿¥ÆÊ£ÑÓŒ¯">
<META HTTP-EQUIV="expires" CONTENT="Mon, 1 Jan 2001 01:01:01 GMT">
<!-- #include file="adovbs.inc" -->
<!--#include file ="../definc_createIISI.asp"-->
</head>
<body onload="document.charset='utf-8';">
<%
if 1 then
   on error resume next
   response.write "   <script>parent.document.all('iframeMapScript').style.display='none';</script>"   
else   
   response.write "   <script>parent.document.all('iframeMapScript').style.display='';</script>"
end if
   
gentime=-1
gisTmpPath="/icormanager/output/"
if request.form("mapAction")<>"" then
   mapAction=CLng(request.form("mapAction"))
elseif request.querystring("mapAction")<>"" then
   mapAction=CLng(request.querystring("mapAction"))
else
   mapAction=2
end if
if request.form("mapTools")<>"" then
   mapTools=CLng(request.form("mapTools"))
elseif request.querystring("mapTools")<>"" then
   mapTools=CLng(request.querystring("mapTools"))
else
   mapTools=2
end if
if request.form("mapExtents")<>"" then
   mapExtents=request.form("mapExtents")
elseif request.querystring("mapExtents")<>"" then
   mapExtents=request.querystring("mapExtents")
else
   mapExtents=""
end if
if request.form("mapScale")<>"" then
   mapScale=request.form("mapScale")
elseif request.querystring("mapScale")<>"" then
   mapScale=request.querystring("mapScale")
else
   mapScale=""
end if
if request.form("mapZoom")<>"" then
   mapZoom=request.form("mapZoom")
elseif request.querystring("mapZoom")<>"" then
   mapZoom=request.querystring("mapZoom")
else
   mapZoom=""
end if
if request.form("mapPx")<>"" then
   mapPx=request.form("mapPx")
elseif request.querystring("mapPx")<>"" then
   mapPx=request.querystring("mapPx")
else
   mapPx=""
end if
if request.form("mapPy")<>"" then
   mapPy=request.form("mapPy")
elseif request.querystring("mapPy")<>"" then
   mapPy=request.querystring("mapPy")
else
   mapPy=""
end if
if request.form("qLayer")<>"" then
   qLayer=request.form("qLayer")
elseif request.querystring("qLayer")<>"" then
   qLayer=request.querystring("qLayer")
else
   qLayer=""
end if
if request.form("qField")<>"" then
   qField=request.form("qField")
elseif request.querystring("qField")<>"" then
   qField=request.querystring("qField")
else
   qField=""
end if
if request.form("qValue")<>"" then
   qValue=request.form("qValue")
elseif request.querystring("qValue")<>"" then
   qValue=request.querystring("qValue")
else
   qValue=""
end if
if request.form("mapPicKind")<>"" then
   mapPicKind=CLng(request.form("mapPicKind"))
elseif request.querystring("mapPicKind")<>"" then
   mapPicKind=CLng(request.querystring("mapPicKind"))
else
   mapPicKind=1
end if
if request.form("mapDynamicZoom")<>"" then
   mapDynamicZoom=CLng(request.form("mapDynamicZoom"))
elseif request.querystring("mapDynamicZoom")<>"" then
   mapDynamicZoom=CLng(request.querystring("mapDynamicZoom"))
else
   mapDynamicZoom=-1
end if
if request.form("mapShowAll")<>"" then
   mapShowAll=CLng(request.form("mapShowAll"))
elseif request.querystring("mapShowAll")<>"" then
   mapShowAll=CLng(request.querystring("mapShowAll"))
else
   mapShowAll=0
end if
if request.form("mapPan")<>"" then
   mapPan=request.form("mapPan")
elseif request.querystring("mapPan")<>"" then
   mapPan=request.querystring("mapPan")
else
   mapPan=""
end if
if request.form("mapLayers")<>"" then
   mapLayers=request.form("mapLayers")
elseif request.querystring("mapLayers")<>"" then
   mapLayers=request.querystring("mapLayers")
else
   mapLayers=""
end if
if request.form("mapLayersStat")<>"" then
   mapLayersStat=request.form("mapLayersStat")
elseif request.querystring("mapLayersStat")<>"" then
   mapLayersStat=request.querystring("mapLayersStat")
else
   mapLayersStat=""
end if
if request.form("mapThema")<>"" then
   mapThema=request.form("mapThema")
elseif request.querystring("mapThema")<>"" then
   mapThema=request.querystring("mapThema")
else
   mapThema="default"
end if
if request.form("mapSize")<>"" then
   mapSize=request.form("mapSize")
elseif request.querystring("mapSize")<>"" then
   mapSize=request.querystring("mapSize")
else
   mapSize=""
end if

response.write "Querystring:<br>" & request.querystring & "<hr>"
response.write "Form:<br>" & request.form & "<hr>"

if mapPicKind=2 or mapDynamicZoom>=0 or mapshowall=1 or mappan<>"" then
   mapActionOld=mapAction
   mapAction=0
end if

amapstr=mapThema & chr(255) & mapAction & chr(255) & mapExtents & chr(255) & cstr(mapTools) & chr(255) & "2" & chr(255) & CStr(mapPicKind) & chr(255) & CStr(mapDynamicZoom) & chr(255) & CStr(mapShowAll) & chr(255) & CStr(mapPan) & chr(255) & CStr(mapLayers) & chr(255) & CStr(mapLayersStat) & chr(255) & CStr(mapSize)
response.write "<script langunage='javascript'>"




if mapAction<=1 then
   select case mapAction
       case 0:   
         if mapDynamicZoom>=0 or mapshowall=1 or mappan<>"" then
            aqstr=""
         else
            aqstr=mapPx & chr(255) & mapPy
         end if      
       case 1:
         if qLayer="Dzia³ki" then
            mapAction=3
            mapTools=3
            response.write "   parent.document.all('MapTools4').checked=true;"
         else
            mapAction=2
            mapTools=2
            response.write "   parent.document.all('MapTools3').checked=true;"
         end if
         mapActionOld=mapAction         
         aqstr=qLayer & chr(255) & qField & chr(255) & qValue
   end select
   astr=amapstr & chr(254) & aqstr
   
   response.write "</script>"
   response.write astr
   ret=Proc_ExecuteMethodRetS("py21","CLASSES_Library_Test_IIS_GeoPlatnikMapa1", "afield", "0", astr, "")
   response.write "<hr>" & ret
   response.write "<script langunage='javascript'>"
   if ret="NORESULTS" then
      response.write "alert('Brak danych!');"
   elseif left(ret,2)="OK" then
      aarrayp0=split(ret,chr(254))
      aarraym=split(aarrayp0(0),chr(255))
      mapCenterX=CDbl(aarraym(1))
      mapCenterY=CDbl(aarraym(2))
      mapScale=int(aarraym(3))
      mapZoomX=CLng(aarraym(4))  'width of a map
      mapZoomY=CLng(aarraym(5))  'heigh of a map
      picWidth=aarraym(6)
      picHeight=aarraym(7)
      mapExtents=aarraym(8)
      randId=aarraym(9)
      mapUnits=aarraym(10)
      picExt=aarraym(11)
      gentime=CDbl(aarrayp0(2))
      response.write "</script>"

      response.write       "<hr>mapCenterX:" &  mapCenterX & "<br>"
      response.write       "mapCenterY" & mapCenterY & "<br>"
      response.write       "mapScale" &   mapScale   & "<br>"
      response.write       "mapZoomX" &   mapZoomX   & "<br>"
      response.write       "mapZoomY" &   mapZoomY   & "<br>"
      response.write       "picWidth" &   picWidth   & "<br>"
      response.write       "picHeight" &  picHeight  & "<br>"
      response.write       "mapExtents" & mapExtents & "<br>"
      response.write       "randid" &     randid     & "<br>"
      response.write       "mapunits" &   mapunits   & "<br>"
      response.write       "pictext" &   pictext   & "<br>"
      response.write       "mapTools" &   mapTools   & "<br>"
      
      
      response.write "<script langunage='javascript'>"   

      aarrayl=split(aarrayp0(1),chr(255))
      aLayersNames=split(aarrayl(0),"|")
      aLayersTypes=split(aarrayl(1),"|")
      aLayersStats=split(aarrayl(2),"|")
      numberLayers=ubound(aLayersNames)
      maprlayers=""
      mapvlayers=""
      dzialki=0
      for i=numberLayers to 0 step -1
         if aLayersNames(i)="Dzia³ki" and aLayersStats(i)=1 then
            dzialki=1
         end if   
         if aLayersStats(i)=0 then
            achecked=""
         else
            achecked="checked"
         end if
         maprlayers1=" <input " & achecked & " class=checkradio type=checkbox id=mapCheckLayers" & CStr(i) & " name=mapCheckLayers" & CStr(i) & " value='" & aLayersNames(i) & "'>&nbsp;" & aLayersNames(i)
         if aLayersTypes(i)="3" then
            maprlayers=maprlayers & maprlayers1
         else
            mapvlayers=mapvlayers & maprlayers1
         end if   
      next
      
      if dzialki=1 then
         response.write "   parent.document.all('geoinfo').style.display='';"
      else
         response.write "   parent.document.all('geoinfo').style.display='none';"
         response.write "   if (parent.document.all('MapTools4').checked==true){parent.document.all('MapTools3').checked=true;};"
      end if
      if maprlayers<>"" then
         maprlayers="<b>Podk³ady rastrowe:</b><br>" & maprlayers
      end if
      response.write "   parent.document.all(""maprlayers"").innerHTML=""" & maprlayers & """;"         
      if mapvlayers<>"" then   
         mapvlayers="<b>Warstwy wektorowe:</b><br>" & mapvlayers
      end if
      response.write "   parent.document.all(""mapvlayers"").innerHTML=""" & mapvlayers & """;"         
      response.write "   parent.numAllLayers=" & numberLayers & ";"
      response.write "   parent.setSearchList();"      
      if mapunits="3" then
         if mapZoomX<=1 then
            zmnoz=100
            sctxt="&nbsp;cm"
         elseif mapZoomX>1 and mapZoomX<=1000 then
            zmnoz=1
            sctxt="&nbsp;m"
         else
            zmnoz=0.001
            sctxt="&nbsp;km"
         end if   
      end if      
      response.write "   parent.document.all('mappicref').src='" & gisTmpPath & "ref_" & randid & ".png';"
      response.write "   parent.document.all('mappicscb').src='" & gisTmpPath & "scb_" & randid & ".png';"
      response.write "   parent.document.all('mappicleg').src='" & gisTmpPath & "leg_" & randid & ".png';"
      response.write "   parent.document.all('mapscale').innerHTML='" & mapScale & "';"
      response.write "   parent.document.all('mapzoom').innerHTML='" & round(mapZoomX*zmnoz,2) & sctxt & "';"
      response.write "   parent.document.all('mappic').width='" & picWidth & "';"
      response.write "   parent.document.all('mappic').height='" & picHeight & "';"
      response.write "   parent.document.all('mappic').src='" & gisTmpPath & "map_" & randid & "." & picext & "';"

            
   end if
elseif mapAction=2 then
   aqstr=mapPx & chr(255) & mapPy
   astr=amapstr & chr(254) & aqstr

   response.write "</script>"
   response.write astr
   ret=Proc_ExecuteMethodRetS("py21","CLASSES_Library_Test_IIS_GeoPlatnikMapa1", "afield", "0", astr, "")   
   response.write "<hr>" & ret & "<br>"
   response.write "<script langunage='javascript'>"
   

   if ret="NORESULTS" then
      response.write "alert('Brak informacji o warstwach');"
   elseif left(ret,2)="OK" then
      aarray0=split(ret,chr(254))
      aarray=split(aarray0(0),",")
      randId=aarray(1)
      gentime=CDbl(aarray0(2))
      axmlfile=gisTmpPath & "inf_" & randid & ".xml"
      response.write "</script>"
         response.write axmlfile & "<br>"
      response.write "<script langunage='javascript'>"
      response.write "   var awindow=window.open('','INFO','scrollbars=yes,toolbar=no,left=10,top=10,width=300,height=500,resizable=yes, status=No, menubar=yes',true);"

      response.write "   awindow.location='" & axmlfile & "';"     
      response.write "   awindow.focus();"
      
   end if



elseif mapAction=3 then
   aqstr=mapPx & chr(255) & mapPy
   astr=amapstr & chr(254) & aqstr
   
   response.write "</script>"
   response.write astr
   ret=Proc_ExecuteMethodRetS("py21","CLASSES_Library_Test_IIS_GeoPlatnikMapa1", "afield", "0", astr, "")   
   response.write "<hr>" & ret
   response.write "<script langunage='javascript'>"
   
   

   if ret="NORESULTS" then
      response.write "alert('Brak informacji o dzia³kach');"
   elseif left(ret,2)="OK" then
      response.write "   var awindow=window.open('','GEOINFO','scrollbars=yes,toolbar=no,left=10,top=10,width=700,height=500,resizable=yes, status=No, menubar=yes',true);"
      aarray0=split(ret,chr(254))
      aarray=split(aarray0(0),",")
      response.write "</script>"      
      nrobrebu=aarray(1)
      nrdzialki=aarray(2)
      gentime=CDbl(aarray0(2))
      response.write       "<hr>nrobrebu:" &  nrobrebu & "<br>"
      response.write       "nrdzialki:" &  nrdzialki & "<br>"      
      response.write "<script langunage='javascript'>"
      ahref="um/wypis.asp?NrDzialki=" & nrdzialki & "&NrObrebu=" & nrobrebu
      response.write "   awindow.location='" & ahref & "';"     
      response.write "   awindow.focus();"
   else
      ret="BAD"
   end if
else
   response.write "   alert('Opcja jeszcze niedostêpna!');"
end if
if ret="BAD" or err.number<>0 then
    response.write "   alert('Problem z dostêpem do serwera mapowego!');"
end if
response.write "   parent.document.all('mappic').style.cursor='crosshair';"
response.write "   parent.document.all('mapdiv').style.display='';"
response.write "   parent.document.all('mapdescr').style.display='none';"
response.write "   parent.document.all('mapThemaList').value='" & mapThema & "';"
response.write "   parent.document.all('mapgentime').innerHTML='" & FormatNumber(gentime,4) & "';"
response.write "   parent.inProgress=false;"
response.write "</script>"
if mapPicKind=2 or mapDynamicZoom>=0 or mapshowall=1 or mappan<>"" then
   mapAction=mapActionOld
end if

%>
<form validate='onchange' invalidColor='gold' mark year4 name=mapScriptForm id=mapScriptForm METHOD='post' ACTION='MapScriptProc.asp'>
  <input type=hidden id='mapAction' name='mapAction' value='<%=mapAction%>'>
  <input type=hidden id='mapExtents' name='mapExtents' value='<%=mapExtents%>'>
  <input type=hidden id='mapScale' name='mapScale' value='<%=mapScale%>'>
  <input type=hidden id='mapZoom' name='mapZoom' value='<%=mapZoom%>'>
  <input type=hidden id='mapPx' name='mapPx' value='<%=mapPx%>'>
  <input type=hidden id='mapPy' name='mapPy' value='<%=mapPy%>'>
  <input type=hidden id='qLayer' name='qLayer' value='<%=qLayer%>'>
  <input type=hidden id='qField' name='qField' value='<%=qField%>'>
  <input type=hidden id='qValue' name='qValue' value='<%=qValue%>'>  
  <input type=hidden id='mapSize' name='mapSize' value='<%=mapSize%>'>  
  <input type=hidden id='mapTools' name='mapTools' value='<%=mapTools%>'>
  <input type=hidden id='mapThema' name='mapThema' value='<%=mapThema%>'>  
  <input type=hidden id='mapLayers' name='mapLayers' value=''>
  <input type=hidden id='mapLayersStat' name='mapLayersStat' value=''>
  <input type=hidden id='mapPicKind' name='mapPicKind' value='1'>
  <input type=hidden id='mapDynamicZoom' name='mapDynamicZoom' value='-1'>
  <input type=hidden id='mapShowAll' name='mapShowAll' value='0'>
  <input type=hidden id='mapPan' name='mapPan' value=''>
</form>
<%

%>
</body>
</html>