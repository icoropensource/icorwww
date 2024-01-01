<%@ codepage=65001 LANGUAGE="VBSCRIPT" %><!--#include file ="definc_createIISI.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<script type="text/javascript" src="/icormanager/inc/xmlextras.js"></script>
<link rel="stylesheet" type="text/css" media="screen" href="superfish.css"> 
<link rel="stylesheet" type="text/css" media="screen" href="superfish-navbar.css"> 
<script type="text/javascript" src="/icorlib/jquery/jquery-latest.min.js"></script>
<link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/ui-1.8.14/css/<%=Dession("UI_SKIN")%>/jquery-ui-1.8.14.custom.min.css">
<script type="text/javascript" src="/icorlib/jquery/plugins/ui-1.8.14/js/jquery-ui-1.8.14.custom.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/hoverIntent/jquery.hoverIntent.min.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/superfish/js/superfish.min.js"></script>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>

<title>ICOR Menu</title>
<style>
   A:link   {text-decoration:none; }
   A:visited {text-decoration:none;}
   A:active {text-decoration:none; }
</style>
<STYLE>
body {
   padding:0px;
   margin:0px;
   font-size:11px;
   font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;
}
/*
a.button {padding:3px;display:inline-block;
          background-image:url('img/tlo_tytul2.gif');background-position:0% 50%;
          border:solid 1px #cccccc;
          text-decoration:none;color:black}
a.button:hover {background-image:url('img/tlo_tytul.gif');
                border:solid 1px #2694e8;}

a {color:blue;text-decoration:none}
a:hover {color:red;}
*/
a img {border:none}

div.user {margin-top:3px;}

.bold {font-weight:bold}
.left {text-align:left}
.right {text-align:right}
.flow_right {float:right}
.flow_left {float:left}
.clear {clear:both}
</STYLE>

<script language="JavaScript" defer>
var lastcols="222,*";
try {
   var tabdict={}; //new ActiveXObject("Scripting.Dictionary");
} catch(ex) {
   var tabdict=null;
   alert("Proszę o kontakt z administratorem - wystąpił problem z konfiguracją ActiveX");
}

var lasttabcid=-1;
var lasttabnr=-1;
var lasttabid=-1;

if (tabdict!=null) {
   tabdict["ICOR"]="ICOR";
}

/*
function toggleMenu()
{
   var s=window.top.document.all("framesetTOCPane").cols;
   if (s=="0,*") {
      s=lastcols;
      if (s=="0,*") s="222,*";
   } else {
      lastcols=s;
      s="0,*";
   }
   window.top.document.all("framesetTOCPane").cols=s;
}
function CallMenuFunction() {
   var menuChoice = event.result;
   var menuhref = event.menuItem.getAttribute("menuhref")
   if (menuhref==null)
      return;
   if (menuhref=="ToggleMenu") {
      toggleMenu();
      return;
   }
   if (menuhref!="") {
      window.top.frames("TEXT").location.href=menuhref;
   }
}
*/

function SetSheetForCid(cid,atoframe,atabid) 
{
   if (tabdict!=null) {
      var lasttab = new Object();
      lasttab.cid=cid;
      lasttab.nr=atoframe;
      lasttab.id=atabid;
      tabdict[cid]=lasttab;
   } else {
      lasttabcid=cid;
      lasttabnr=atoframe;
      lasttabid=atabid;
   }
}

function GetSheetForCid(cid)
{
   if (tabdict!=null) {
      if (cid in tabdict) {
         lasttab=tabdict[cid];
         return lasttab;
      }
   } else {
      if (lasttabcid==cid) {
         var newtab = new Object();
         newtab.cid=lasttabcid;
         newtab.nr=lasttabnr;
         newtab.id=lasttabid;
         return newtab;
      }
   }
}

function displayMessage(s) {
   alert(s);
}

//<%=Dession("uid")%>
function StartProgress(aname) {
   var ProgressURL='progresswindow.asp?name='+aname
   var v=window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')
}

function contextForBody(obj)
{
//   var popupoptions;
//   popupoptions = [
//                  new ContextItem("Odśwież menu",function(){window.location.reload();})
//                  new ContextItem("Disabled Item",null,true),
//                  new ContextSeperator(),
//                  new ContextItem("Print",function(){window.print()},true),
//                  new ContextItem("View Source",function(){location='view-source:'+location}),
//                  new ContextItem("Show Text",function(){alert(obj.tagName);})
//                 ]
//   ContextMenu.display(popupoptions)
   window.event.cancelBubble = true;
   window.event.returnValue  = false;
}

var maxHappySound=6;
var maxSadSound=5;

function playHappySound() {
  var anum=Math.random();
  anum=anum*maxHappySound;
  anum=1+Math.floor(anum);
  try {
     var x=document.all("happySound"+anum);
     x.play();
  } catch(e) {
     document.all("sounds").insertAdjacentHTML("beforeEnd","<embed src='inc/snd/success/happy"+anum+".wav' hidden=true autostart=true loop=false id='happySound"+anum+"' name='happySound"+anum+"' VOLUME=100>");
  }
}

function playSadSound() {
  var anum=Math.random();
  anum=anum*maxSadSound;
  anum=1+Math.floor(anum);
  try {
     var x=document.all("sadSound"+anum);
     x.play();
  } catch(e) {
     document.all("sounds").insertAdjacentHTML("beforeEnd","<embed src='inc/snd/sad/sad"+anum+".wav' hidden=true autostart=true loop=false id='sadSound"+anum+"' name='sadSound"+anum+"' VOLUME=100>");
  }
}

try {
   var tabstate={}; //new ActiveXObject("Scripting.Dictionary");
} catch(ex) {
   var tabstate=null;
   alert("Proszę o kontakt z administratorem - wystąpił problem z konfiguracją ActiveX");
}

function stateChecked(oXmlDoc,astateid,atext) {
   if( oXmlDoc == null || oXmlDoc.documentElement == null) {
      if( oXmlDoc == null) {
         alert("State oXmlDoc (1) == null");
      }
   } else {
      var root = oXmlDoc.documentElement;
      var cs = root.childNodes;
      var l = cs.length;
      var aname="";
      var avalue="";
      for (var i = 0; i < l; i++) {
         if (cs[i].tagName == "STATE") {
            aname=cs[i].getAttribute("name");
            avalue=cs[i].getAttribute("value");
         }
      }
      if (avalue=="DEL") {
         try {
            delete tabstate[astateid];
         } catch(ex) {;}
      }
      if (avalue=="OK") {
        delete tabstate[astateid];
        window.focus();
        playHappySound();
//        document.all("stateinfo").stop();
//        document.all("stateinfo").style.display='none';
        alert(aname);
      }
      if (avalue=="BAD") {
        delete tabstate[astateid];
        window.focus();
        playSadSound();
//        document.all("stateinfo").stop();
//        document.all("stateinfo").style.display='none';
        alert(aname);
      }
      if (avalue.slice(0,1)=='#') {
//        document.all("stateinfo").innerHTML=avalue.slice(1);
//        document.all("stateinfo").style.display='';
//        document.all("stateinfo").start();
      }
//      if (avalue=="RUN") {
//        tabstate.Remove(astateid);
//      }
   }
}

// $$ do zmiany dla FF
function stateChecker() {
  var astateid,w;
  w=0;
  for (astateid in tabstate) {
    w=1;
    var xmlHttp = XmlHttp.create();
    xmlHttp.open("GET","icorsync.asp?mode=get&state="+astateid,true);
    xmlHttp.onreadystatechange=function () {
      if (xmlHttp.readyState == 4) {
        stateChecked(xmlHttp.responseXML,astateid,xmlHttp.responseText);
      }
    };
    try {
      window.setTimeout(function () {
        try {
          xmlHttp.send(null);
        } catch(ex) {;
        }
      }, 10);
    } catch(ex) {;
    }
  }
  if (w==1) {
    window.setTimeout(stateChecker,7000);
  }
}

function registerStateBadOK(aid) {
  tabstate[aid]=aid;
  window.setTimeout(stateChecker,7000);
}

//try {
//  document.all("stateinfo").stop();
//} catch(ex) {;
//}
</script>
<base target="TEXT">
</head>
<body>
<div id="mdiv2" class="ui-widget ui-widget-header">
   <div class="flow_left">
      <a target=new href="http://www.mikroplan.com.pl"><IMG SRC="images/logo_mp2b.gif" class="logo_mp"></a>
   </div>
   <div class="flow_left">&nbsp;&nbsp;&nbsp;
   </div>
   <a name="mymenu"></a>
<%
dim iret
iret=Process("NavigationBar")
%>
<!--#include file ="definc_disposeIISI.asp"-->
   <div class="flow_right right">
      <a target=new href="http://www.icor.pl"><IMG SRC="images/logo_icor2b.gif" class="logo_icor"></a>
      <div class="user">
<%
if Dession("username")="" then
   Dession("username")=ICORExecuteMethod("CLASSES\Library\NetBase\WWW\Server\GetUserByUID", "", -1, "", "")
   '/Dession("uid")
end if
%>
         Jesteś zalogowany jako: <span class="bold"><%=Dession("username")%> (<%=Dession("UI_SKIN")%>)</span> | <a href="logout.asp">Wyloguj</a>
      </div>
   </div>
   <div class="clear"></div>
   <div style="display:none;" id="sounds" name="sounds">
   </div>
</div>
<SCRIPT LANGUAGE=javascript>
jQuery(function() {
   jQuery("body").append("<div id='mdiv1' class='ui-widget ui-widget-content'></div>");
   var el=jQuery("#mdiv1");
   var abc=el.css("background-color");
   jQuery("html").css("background-color",abc);
   jQuery("body").css("background-color",abc);
   el.remove();

   jQuery('ul.sf-menu').superfish();
});
</SCRIPT>
</body></html>
