<%@ CODEPAGE=65001 LANGUAGE="VBSCRIPT" %><!--#include file ="definc_createIISI.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
SHOW_TABS=0
MENU_CLASSIC=0
MENU_FAVORITES=0
MENU_JSTREE=1
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<title>Menu</title>

<%
if MENU_CLASSIC=1 then
%>
<script type="text/javascript" src="/icormanager/inc/xtree.js"></script>
<script type="text/javascript" src="/icormanager/inc/xmlextras.js"></script>
<script type="text/javascript" src="/icormanager/inc/xloadtree.js"></script> 
<script type="text/javascript" src="/icormanager/inc/ContextMenu.js"></script>
<%
end if
%>
<script type="text/javascript" src="/icorlib/jquery/jquery-1.3.2.js"></script>
<link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/ui-1.7.2/css/<%=Dession("UI_SKIN")%>/jquery-ui-1.7.2.custom.min.css">
<script type="text/javascript" src="/icorlib/jquery/plugins/ui-1.7.2/js/jquery-ui-1.7.2.custom.js"></script>

<%
if MENU_FAVORITES=1 or MENU_JSTREE=1 then
%>
<!--
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/lib/jquery.cookie.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/lib/jquery.hotkeys.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/lib/jquery.metadata.js"></script>
-->

<!--
$$<script type="text/javascript" src="#/icorlib/jquery/plugins/jstree/source/lib/sarissa.js"></script>
-->

<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/jquery.tree.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.contextmenu.js"></script>
<!--
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.checkbox.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.cookie.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.hotkeys.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.metadata.js"></script>
<script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.themeroller.js"></script>
-->
<!--
$$  <script type="text/javascript" src="#/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.xml_flat.js"></script>
$$  <script type="text/javascript" src="#/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.xml_nested.js"></script>
-->
<link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/jstree/source/themes/default/style.css">
<script type="text/javascript" src="/icorlib/json/json2.js"></script>
<%
end if
%>

<STYLE>
body, html { width: 100%; height: 100%; }

body {
/*
   background-color: #fbfbfb;
*/
   margin:0px;padding:0px;
   font-size:12px;font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;
}
a img {border:none}
div.menu {padding:2px}

<%
if MENU_CLASSIC=1 then
%>
A:link {
   COLOR: black; TEXT-DECORATION: none
}
A:visited {
   COLOR: black; TEXT-DECORATION: none
}
A:active {
   COLOR: black; TEXT-DECORATION: none
}
.webfx-tree-container {
   PADDING-BOTTOM: 0px; MARGIN: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; FONT: icon; WHITE-SPACE: nowrap; PADDING-TOP: 0px
}
.webfx-tree-item {
   PADDING-BOTTOM: 0px; MARGIN: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; FONT: icon; WHITE-SPACE: nowrap; COLOR: black; PADDING-TOP: 0px
}
.webfx-tree-item A {
   PADDING-BOTTOM: 1px; PADDING-LEFT: 2px; PADDING-RIGHT: 2px; MARGIN-LEFT: 3px; PADDING-TOP: 1px
}
.webfx-tree-item A:active {
   PADDING-BOTTOM: 1px; PADDING-LEFT: 2px; PADDING-RIGHT: 2px; MARGIN-LEFT: 3px; PADDING-TOP: 1px
}
.webfx-tree-item A:hover {
   PADDING-BOTTOM: 1px; PADDING-LEFT: 2px; PADDING-RIGHT: 2px; MARGIN-LEFT: 3px; PADDING-TOP: 1px
}
.webfx-tree-item A {
   COLOR: black; TEXT-DECORATION: none
}
.webfx-tree-item A:hover {
   COLOR: blue; TEXT-DECORATION: underline
}
.webfx-tree-item A:active {
   BACKGROUND: #e1a9a8; TEXT-DECORATION: none
}
.webfx-tree-item IMG {
   BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; VERTICAL-ALIGN: middle; BORDER-LEFT-WIDTH: 0px
}
.webfx-tree-icon {
   WIDTH: 16px; HEIGHT: 16px
}
.webfx-tree-item A.selected {
   BACKGROUND: #e1a9a8
}
.webfx-tree-item A.selected-inactive {
   BACKGROUND: #e1a9a8
}
INPUT.tree-check-box {
   PADDING-BOTTOM: 0px; MARGIN: 0px; PADDING-LEFT: 0px; WIDTH: auto; PADDING-RIGHT: 0px; HEIGHT: 14px; VERTICAL-ALIGN: middle; PADDING-TOP: 0px
}
.WebFX-ContextMenu {
   BORDER-RIGHT-WIDTH: 0px; WIDTH: 10px; BORDER-TOP-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px
}
.WebFX-ContextMenu-Body {
   BORDER-BOTTOM: 2px outset; BORDER-LEFT: 2px outset; PADDING-BOTTOM: 1px; BACKGROUND-COLOR: menu; PADDING-LEFT: 1px; PADDING-RIGHT: 1px; BORDER-TOP: 2px outset; BORDER-RIGHT: 2px outset; PADDING-TOP: 1px; margin: 0px
}
.WebFX-ContextMenu-Item {
   PADDING-BOTTOM: 2px; PADDING-LEFT: 16px; WIDTH: 100%; PADDING-RIGHT: 16px; FONT: menu; COLOR: menutext; CURSOR: default; PADDING-TOP: 2px
}
.WebFX-ContextMenu-Over {
   PADDING-BOTTOM: 2px; BACKGROUND-COLOR: highlight; PADDING-LEFT: 16px; WIDTH: 100%; PADDING-RIGHT: 16px; FONT: menu; COLOR: highlighttext; CURSOR: default; PADDING-TOP: 2px
}
.WebFX-ContextMenu-Disabled {
   PADDING-BOTTOM: 2px; PADDING-LEFT: 16px; WIDTH: 100%; PADDING-RIGHT: 16px; FONT: menu; COLOR: graytext; CURSOR: default; PADDING-TOP: 2px
}
.WebFX-ContextMenu-Disabled-Over {
   PADDING-BOTTOM: 2px; BACKGROUND-COLOR: highlight; PADDING-LEFT: 16px; WIDTH: 100%; PADDING-RIGHT: 16px; FONT: menu; COLOR: graytext; CURSOR: default; PADDING-TOP: 2px
}
.WebFX-ContextMenu-Separator {
   BORDER-BOTTOM: 1px inset; BORDER-LEFT: 1px inset; MARGIN: 3px 1px; HEIGHT: 2px; FONT-SIZE: 0pt; OVERFLOW: hidden; BORDER-TOP: 1px inset; BORDER-RIGHT: 1px inset
}
.WebFX-ContextMenu-Disabled-Over .WebFX-ContextMenu-DisabledContainer {
   WIDTH: 100%; DISPLAY: block
}
.WebFX-ContextMenu-Disabled-Over .WebFX-ContextMenu-DisabledContainer .WebFX-ContextMenu-DisabledContainer {
   
}
#main {
   POSITION: relative; WIDTH: 100%; HEIGHT: 100%; TOP: 0%; LEFT: 0%
}
.label2 {
   BORDER-BOTTOM: 0px dotted; BORDER-LEFT: 0px dotted; FONT-FAMILY: arial, verdana; BACKGROUND: #6699cc; COLOR: #ffffff; FONT-SIZE: 12px; BORDER-TOP: #f1f1f1 1px solid; CURSOR: pointer; FONT-WEIGHT: normal; BORDER-RIGHT: 0px dotted
}
.label2 :hover {
   CURSOR: pointer
}
.accBody {
   BORDER-BOTTOM: #84a3d1 0px solid; BORDER-LEFT: #84a3d1 0px solid; BACKGROUND: red; OVERFLOW: scroll; BORDER-TOP: #84a3d1 0px solid; BORDER-RIGHT: #84a3d1 0px solid
}
<%
end if
%>
</STYLE>
<%
if MENU_CLASSIC=1 then
%>
<!--[if IE]>
<STYLE>
.WebFX-ContextMenu-Disabled .WebFX-ContextMenu-DisabledContainer {
   FILTER: chroma(color=#010101) dropshadow(color=ButtonHighlight, offx=1, offy=1); WIDTH: 100%; DISPLAY: block; BACKGROUND: graytext
}
.WebFX-ContextMenu-Disabled .WebFX-ContextMenu-DisabledContainer .WebFX-ContextMenu-DisabledContainer {
   FILTER: gray() chroma(color=#ffffff) chroma(color=#fefefe) chroma(color=#fdfdfd) chroma(color=#fcfcfc) chroma(color=#fbfbfb) chroma(color=#fafafa) chroma(color=#f9f9f9) chroma(color=#f8f8f8) chroma(color=#f7f7f7) chroma(color=#f6f6f6) chroma(color=#f5f5f5) chroma(color=#f4f4f4) chroma(color=#f3f3f3) mask(color=#010101); BACKGROUND: none transparent scroll repeat 0% 0%
}
</STYLE>
<![endif]-->
<%
end if
%>
<base target="TEXT">
</head>
<%
if MENU_CLASSIC=1 then
%>
<body onload="ContextMenu.intializeContextMenu()" >
<script language="javascript">
//oncontextmenu="contextForBody(this)"
function contextForBody(obj)
{
   var popupoptions;
   popupoptions = [
                  new ContextItem("Odśwież menu",function(){window.location.reload();})
//                  new ContextItem("Disabled Item",null,true),
//                  new ContextSeperator(),
//                  new ContextItem("Print",function(){window.print()},true),
//                  new ContextItem("View Source",function(){location='view-source:'+location}),
//                  new ContextItem("Show Text",function(){alert(obj.tagName);})
                 ]
   ContextMenu.display(popupoptions)
}

var copiedNode=null;
function processMenuAction(action) {
   if (action=="") {
      return;
   }
   var xmlHttp = XmlHttp.create();
   xmlHttp.open("GET", action, false);
   xmlHttp.send(null);
   oXmlDoc=xmlHttp.responseXML;
   if( oXmlDoc == null || oXmlDoc.documentElement == null) {
      return
   } else {
      var root = oXmlDoc.documentElement;
      var cs = root.childNodes;
      var l = cs.length;
      for (var i = 0; i < l; i++) {
         if (cs[i].tagName == "item") {
            var text = getAsPLText(cs[i].getAttribute("text"));
            var action=cs[i].getAttribute("action")
            var value=cs[i].getAttribute("value")
            if (action=='redirect') {
               window.top.frames("TEXT").location.href=value;
            }
         }
      }
   }
}
function myContextMenuHandler(obj,event) {
   var anode=tree.getSelected();
   var popupoptions;
   var w=0;
   popupoptions = [
//      new ContextItem("Item: "+obj.tagName+' / '+anode.id,null,true)
   ]
   if (anode.context) {
      var xmlHttp = XmlHttp.create();
      xmlHttp.open("GET", anode.context, false);
      xmlHttp.send(null);
      oXmlDoc=xmlHttp.responseXML;
      if( oXmlDoc == null || oXmlDoc.documentElement == null) {
         alert ('oxmldoc=null');
         return
      } else {
         var root = oXmlDoc.documentElement;
         var cs = root.childNodes;
         var l = cs.length;
         for (var i = 0; i < l; i++) {
            if (cs[i].tagName == "menuitem") {
               var text = getAsPLText(cs[i].getAttribute("text"));
               if (text=="-") {
                  popupoptions.push(new ContextSeperator());
                  w=0;
               } else {
                  popupoptions.push(new ContextItem(text,function(aparam){processMenuAction(aparam);},false,cs[i].getAttribute("action")));
                  w=1;
               }
            }
         }
      }
   }
   if (anode.allowmove=="1") {
      popupoptions.push(new ContextSeperator());
      popupoptions.push(new ContextItem("Przenieś rozdział",function(){contextCopyNode(obj);}));
      if (copiedNode!=null) {
         popupoptions.push(new ContextItem("Wstaw jako podrozdział",function(){contextPasteNode(obj,"0");}));
         popupoptions.push(new ContextItem("Wstaw przed rozdziałem",function(){contextPasteNode(obj,"1");}));
      }
   }
   if ((popupoptions.length)>0 && (w==1)) {
      popupoptions.push(new ContextSeperator());
   }
<%
   if MENU_FAVORITES=1 then
%>
   popupoptions.push(new ContextItem("Dodaj do ulubionych",function(){contextFavoritesNode(obj);}));
<%
   end if
%>
   popupoptions.push(new ContextItem("Odśwież rozdział",function(){contextRefreshNode(obj);}));
   popupoptions.push(new ContextItem("Odśwież menu",function(){window.location.reload();}));
   ContextMenu.display(popupoptions,event);
}
function contextRefreshNode(obj) {
   var anode=tree.getSelected();
   anode.doRefresh();
}

function contextCopyNode(obj) {
   var anode=tree.getSelected();
   copiedNode=anode;
}
function contextPasteNode(obj,aparam) {
   if (copiedNode==null) {
      return false;
   }
   if (copiedNode.coid==null) {
      return false;
   }
   var anode=tree.getSelected();
   if (anode.coid==null) {
      return false;
   }
   var xmlHttp = XmlHttp.create();
   var action="icormain.asp?jobtype=workflowmenustructchaptermove&coid1="+copiedNode.coid+"&coid2="+anode.coid+"&param="+aparam+"&XMLData=1"
   xmlHttp.open("GET", action, false);
   xmlHttp.send(null);
   oXmlDoc=xmlHttp.responseXML;
   if( oXmlDoc == null || oXmlDoc.documentElement == null) {
      return
   } else {
      var root = oXmlDoc.documentElement;
      var cs = root.childNodes;
      var l = cs.length;
      for (var i = 0; i < l; i++) {
         if (cs[i].tagName == "item") {
            var text = getAsPLText(cs[i].getAttribute("text"));
            var action=cs[i].getAttribute("action");
            if (action=="OK") {
               copiedNode.remove();
               if (anode.parentNode==null) {
                  tree.reload();
               } else {
                  anode.parentNode.doRefresh();
               }
               copiedNode=null;
            }
            if (action=="BAD") {
               alert("BAD "+text);
            }
         }
      }
   }
   return false;
}
</script>
<%
else
%>
<body>
<%
end if
%>
<%
if MENU_FAVORITES=1 then
%>
<script language="javascript">
function addFavoritesNode(anode,btree,pnode) {
   var astate="open";
   var anode2;
   if (anode.src) {
      astate="closed";
   }
   var bnode=btree.create({
      data:{title:anode.text,state:astate,icon:anode.icon,attributes:{href:anode.action,menucontext:anode.context,openicon:anode.openIcon,isfav:"1",allowmove:anode.allowmove,coid:anode.coid,src:anode.src}}
      },
      pnode,"inside"
   );
   if (anode.src) {
      return;
   }
   for (var i=0;i<anode.childNodes.length;i++) {
      anode2=anode.childNodes[i];
      addFavoritesNode(anode2,btree,bnode);
   }
}

function contextFavoritesNode(obj) {
   var anode=tree.getSelected();
   var btree=jQuery.tree.focused();
   addFavoritesNode(anode,btree,-1);
}
</script>
<%
end if
%>
<!--
<div id=mainHeaderButtons></div>
<script type="text/javascript" src="http://ui.jquery.com/applications/themeroller/themeswitchertool/"></script>
<script type="text/javascript"> jQuery(function(){ jQuery('<div style="position:absolute;right:20px;"></div>').insertAfter('#mainHeaderButtons').themeswitcher({height:220,buttonPreText:'Sk�rka:',initialText:'Zmiana sk�rki'});});</script>
-->
<%
if SHOW_TABS=1 then
%>
<div id="tabscontainer-1">
   <ul>
<%
   if MENU_CLASSIC=1 then
%>
      <li><a href="#tabs-1">Menu</a></li>
<%
   end if
   if MENU_JSTREE=1 then
%>
      <li><a href="#tabs-2">Menu JS</a></li>
<%
   end if
%>
<!--
      <li><a href="#tabs-3">Konfiguracja</a></li>
-->      
<%
   if MENU_FAVORITES=1 then
%>
      <li><a href="#tabs-4">Ulubione</a></li>
<%
   end if
%>
   </ul>
   <div id="tabs-1">
<%
end if
%>
<%
if MENU_CLASSIC=1 then
%>
<script type="text/javascript">
webFXTreeHandler.oncontextmenu=myContextMenuHandler;
var tree=new WebFXLoadTree("ICOR", "icormain.asp?jobtype=menuxml&OID=-1&XMLData=1");
//tree.setBehavior("classic");
document.write(tree);
tree.reload();
</script>
<%
end if
if SHOW_TABS=1 then
%>
   </div>
<%
end if
%>
<%
if SHOW_TABS=1 then
%>
   <div id="tabs-2">
<%
end if
if MENU_JSTREE=1 then
%>
      <div class="demo" id="tree_menu">
      </div>
<%
end if
if SHOW_TABS=1 then
%>
   </div>
<%
end if
%>
<!--
   <div id="tabs-3">
      <p>Mauris eleifend est et turpis. Duis id erat. Suspendisse potenti. Aliquam vulputate, pede vel vehicula accumsan, mi neque rutrum erat, eu congue orci lorem eget lorem. Vestibulum non ante. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Fusce sodales. Quisque eu urna vel enim commodo pellentesque. Praesent eu risus hendrerit ligula tempus pretium. Curabitur lorem enim, pretium nec, feugiat nec, luctus a, lacus.</p>
      <p>Duis cursus. Maecenas ligula eros, blandit nec, pharetra at, semper at, magna. Nullam ac lacus. Nulla facilisi. Praesent viverra justo vitae neque. Praesent blandit adipiscing velit. Suspendisse potenti. Donec mattis, pede vel pharetra blandit, magna ligula faucibus eros, id euismod lacus dolor eget odio. Nam scelerisque. Donec non libero sed nulla mattis commodo. Ut sagittis. Donec nisi lectus, feugiat porttitor, tempor ac, tempor vitae, pede. Aenean vehicula velit eu tellus interdum rutrum. Maecenas commodo. Pellentesque nec elit. Fusce in lacus. Vivamus a libero vitae lectus hendrerit hendrerit.</p>
   </div>
-->
<%
if SHOW_TABS=1 then
%>
   <div id="tabs-4">
<%
end if
if MENU_FAVORITES=1 then
%>

      <div class="demo" id="tree_fav">
<!--
         <ul>
            <li id="phtml_1" class="open"><a href="#"><ins>&nbsp;</ins>Root node 1</a>
               <ul>
                  <li id="phtml_2"><a href="#"><ins>&nbsp;</ins>Child node 1</a></li>
                  <li id="phtml_3"><a href="#"><ins>&nbsp;</ins>Child node 2</a></li>
                  <li id="phtml_4"><a href="#"><ins>&nbsp;</ins>Some other child node with longer text</a></li>
               </ul>
            </li>
            <li id="phtml_5"><a href="#"><ins>&nbsp;</ins>Root node 2</a></li>
            <li id="phtml_5"><a href="#"><ins>&nbsp;</ins>Root node 3</a></li>
            <li id="phtml_5"><a href="#"><ins>&nbsp;</ins>Root node 4</a></li>
         </ul>
-->
      </div>
<%
end if
%>
<%
if SHOW_TABS=1 then
%>
   </div>
</div>
<%
end if
%>

<%
   if MENU_FAVORITES=1 then
%>
<script type="text/javascript">
function saveTree() {
   var atree=jQuery.tree.focused();
   var atext=atree.get([],"json",{
      outer_attrib : [ "id", "rel", "class"],
      inner_attrib : [ "href","menucontext","openicon","isfav"]
   });
   atext=JSON.stringify(atext);
//   alert(atext);
//   return
//   var atext=t.get([],"html");
   jQuery.ajax({
      url: "icormain.asp?jobtype=menuxml&OID=-1&XMLData=json",
      async: false,
      cache: false,
      type: "POST",
      data: ({data : atext}),
      dataType: "html"
   });
}

jQuery(function () { 
   jQuery("#tree_fav").tree({
      data : {
         type : "json",
         async : true,
         opts : {
            async: true,
            url : "icormain.asp?jobtype=menufavget"
         }
      },
      plugins:{ 
         contextmenu:{

            items : {
               create : {
                  label : "Nowy rozdział" //,
/*
                  icon  : "create",
                  visible  : function (NODE, TREE_OBJ) { 
                     if(NODE.length != 1) return 0; 
                     return TREE_OBJ.check("creatable", NODE); 
                  }, 
                  action   : function (NODE, TREE_OBJ) { 
                     TREE_OBJ.create(false, TREE_OBJ.get_node(NODE[0])); 
                  },
                  separator_after : true
*/                  
               },
               rename : {
                  label : "Zmień nazwę" //, 
/*
                  icon  : "rename",
                  visible  : function (NODE, TREE_OBJ) { 
                     if(NODE.length != 1) return false; 
                     return TREE_OBJ.check("renameable", NODE); 
                  }, 
                  action   : function (NODE, TREE_OBJ) { 
                     TREE_OBJ.rename(NODE); 
                  } 
*/                  
               },
               remove : {
                  label : "Usuń" //,
/*
                  icon  : "remove",
                  visible  : function (NODE, TREE_OBJ) { 
                     var ok = true; 
                     $.each(NODE, function () { 
                        if(TREE_OBJ.check("deletable", this) == false) {
                           ok = false; 
                           return false; 
                        }
                     }); 
                     return ok; 
                  }, 
                  action   : function (NODE, TREE_OBJ) { 
                     $.each(NODE, function () { 
                        TREE_OBJ.remove(this); 
                     }); 
                  } 
*/                  
               }
            }

         }
      },
      callback : {
         onload : function(TREE_OBJ) {
//            jQuery("a[src]").each(function(i){
//               alert(this.id);
//            });
//            alert('onload');
         },
         onselect : function (n,t) {
            var anode=t.get_node(n);
            var aanchor=jQuery("a",anode);
            var ahref=aanchor.attr("href");
            if (ahref!="#" && ahref!="") {
               window.top.frames("TEXT").location.href=ahref;
            }
         },
         onrename: function(NODE, TREE_OBJ, RB) {
            window.setTimeout(saveTree,200);
         },
         onmove: function(NODE, REF_NODE, TYPE, TREE_OBJ, RB) {
            window.setTimeout(saveTree,200);
         },
         oncreate: function(NODE, REF_NODE, TYPE, TREE_OBJ, RB) {
            window.setTimeout(saveTree,200);
         },
         ondelete: function(NODE, TREE_OBJ, RB) {
            window.setTimeout(saveTree,200);
         },
         ondrop: function(NODE, REF_NODE, TYPE, TREE_OBJ) {
            window.setTimeout(saveTree,200);
         },
         beforedata : function (NODE, TREE_OBJ) {
            alert('beforedata');
            var ret={};
            if (NODE) {
               ret.id=$(NODE).attr("id");
               ret.src=$(NODE).attr("src");
            }
            return ret;
         }
      }
   });
});

/*
mod : {
                                               label   : { "en" : "Modify", "fr" : "Modifier" },
                                               icon    : "menu-mod", // you can set this to a classname or a path
to an icon like ./ico_modif.gif
                                               visible : function (NODE, TREE_OBJ) {
                                                       // this action will be disabled if more than one node is
selected
                                                       if(NODE.length != 1) return 0;
                                                       // otherwise - OK
                                                       return 1;
                                               },
                                               action  : function (NODE, TREE_OBJ) {
                                                       document.location.href = $(NODE).children("a:eq(2)").attr
("href");
                                               },
                                               separator_before : true
                                       }
*/
</script>
<%
end if
%>
<%
if MENU_JSTREE=1 then
%>
<script type="text/javascript">

function processRedirect(ahref) {
   var apos=ahref.indexOf('!http');
   if (apos>=0) {
      ahref=ahref.substr(apos+1);
      window.open(ahref);
      return;
   }
   if (ahref.substr(0,1)=="!") {
      ahref=ahref.substr(1);
      window.open(ahref);
   } else {
      window.top.frames["TEXT"].document.location.href=ahref;
   }
}

page_params={
   current_attachment_node:"",
   current_attachment_coid:"",
   current_attachment_roid:""
}

function TreeNodeRefresh(NODE,TREE_OBJ) {
   var w=0,anode,bnode;
   try {
      var anode=TREE_OBJ.parent(NODE);
      if (anode) {
         TREE_OBJ.refresh(anode);
      } else {
         anode=NODE;
         TREE_OBJ.refresh(anode);
      }
   } catch (e) { 
      w=1;
   };
   if (w==1) {
      try {
         var bnode=TREE_OBJ.parent(anode);
         if (bnode) {
            TREE_OBJ.refresh(bnode);
         }
      } catch (e) { 
      };
   }
}

jQuery(function () { 
   jQuery("#tree_menu").tree({
      data : {
         type : "json",
         async : true,
         opts : {
            async: true,
            url : "icormain.asp?jobtype=menuxml&OID=-1&XMLData=json"
         }
      },
      ui : {
         animation : 0,
         theme_name : 'icor' //'themeroller'
      },
      lang : {
         new_node : "Nowa pozycja",
         loading     : "Ładowanie..."
      },
      types : {
         "menustruct" : {
            clickable   : true, // can be function
            renameable  : false, // can be function
            deletable   : false, // can be function
            creatable   : false, // can be function
            draggable   : false, // can be function
            valid_children : ["chapter"] // all, none, array of values can be function
         },
         "chapter" : {
            clickable   : true, // can be function
            renameable  : true, // can be function
            deletable   : false, // can be function
            creatable   : false, // can be function
            draggable   : true, // can be function
            valid_children : ["chapter","attachment"] // all, none, array of values can be function
         },
         "attachment" : {
            clickable   : true, // can be function
            renameable  : false, //true, // can be function
            deletable   : false, // can be function
            creatable   : false, // can be function
            draggable   : false, //true, // can be function
            valid_children : ["attachment"] // all, none, array of values can be function
         },
         "chaptergroup" : {
            clickable   : true, // can be function
            renameable  : true, // can be function
            deletable   : false, // can be function
            creatable   : false, // can be function
            draggable   : false, // can be function
            valid_children : ["chapter"] // all, none, array of values can be function
         },
         "section" : {
            clickable   : true, // can be function
            renameable  : false, // can be function
            deletable   : false, // can be function
            creatable   : false, // can be function
            draggable   : false, // can be function
            valid_children : ["section"] // all, none, array of values can be function
         }
      },
      plugins:{ 
         contextmenu:{
            items : {
               create : {
                  label : "Nowy rozdział",
                  visible : function (NODE, TREE_OBJ) {
                     return false;
                  }
               },
               rename : {
                  label : "Zmień nazwę", 
                  visible : function (NODE, TREE_OBJ) {
                     return false;
                  }
               },
               remove : {
                  label : "Usuń",
                  visible : function (NODE, TREE_OBJ) {
                     return false;
                  }
               }
            }

         }
      },
      callback : {
         onselect : function (NODE, TREE_OBJ) {
            var anode=TREE_OBJ.get_node(NODE);
            var aanchor=jQuery("a",anode);
            var ahref=aanchor.attr("href");
            if (ahref!="#" && ahref!="") {
               processRedirect(ahref);
            }
         },
         beforeopen : function(NODE, TREE_OBJ) {
            var anode=TREE_OBJ.get_node(NODE);
            var aanchor=jQuery("a",anode);
            var aurl=aanchor.attr("submenu");
            if (aurl=="") {
               return false;
            }
            TREE_OBJ.settings.data.opts.url=aurl;
            return true
         },
         beforemove : function(NODE, REF_NODE, TYPE, TREE_OBJ) {
            var anode=TREE_OBJ.get_node(NODE);
            var arel=jQuery(anode).attr("rel");
            var coid=jQuery("a",anode).attr("coid");
            
            var anoderef=TREE_OBJ.get_node(REF_NODE);
            var arelref=jQuery(anoderef).attr("rel");
            var coidref=jQuery("a",anoderef).attr("coid");
            
            if ((coid=="") || (coidref=="")) {
               return false;
            }
//            if (((TYPE=='before') || (TYPE=='after')) && (arelref=='chaptergroup')) {
//               return false;
//            }
            if ((arelref=='menustruct') && ((TYPE=='before') || (TYPE=='after'))) {
               return false;
            }
//            alert(TREE_OBJ.get_text(NODE)+' -> '+TREE_OBJ.get_text(REF_NODE)+' |  '+arel+' | '+arelref+' | '+TYPE);
            var res=false;
            if (1==1) {
               jQuery.ajax({
                  url: "icormain.asp?jobtype=workflowmenustructchaptermovedrag&coid1="+coid+"&coid2="+coidref+"&type="+TYPE+"&rel1="+arel+"&rel2="+arelref+"&XMLData=json",
                  async: false,
                  cache: false,
                  type: "GET",
                  dataType: "json",
                  success: function (data, textStatus) {
                     if (textStatus=='success') {
                        if (data.status!='OK') {
                           alert(data.info);
                        } else {
                           res=true;
                        }
                     } else {
                        alert("Problem z przeniesieniem rozdziału.");
                     };
                  }
               });
            }
            return res;
         },
/*
         onrgtclk : function(NODE,TREE_OBJ,EV) { 
            EV.preventDefault();
            EV.stopPropagation();
            return false;
         },
 */
         onrgtclkbefore : function(NODE, TREE_OBJ, EV) {
            jQuery.tree.plugins.contextmenu.defaults.items={};
            TREE_OBJ.settings.plugins.contextmenu.items={};
            var amenucontext=jQuery(NODE).find('> a').attr('context');
            var amenushow=false;
            var litems={};
            var i=0;
            jQuery.ajax({
               url: amenucontext,
               async: false,
               cache: false,
               type: "GET",
               dataType: "json",
               success: function (data, textStatus) {
                  if (textStatus=='success') {
                     var wseparator=false;
                     for (i in data) {
                        if (data[i].text!='-') {
                           amenushow=true;
                           litems['item_'+i.toString()]={
                              label:data[i].text,
                              itemaction:data[i].action,
                              icon:data[i].icon,
                              separator_before:wseparator,
                              action: function (NODE,TREE_OBJ,ITEM) {
                                 jQuery.ajax({
                                    url: ITEM.itemaction,
                                    async: false,
                                    cache: false,
                                    type: "GET",
                                    dataType: "json",
                                    success: function (rdata, textStatus) {
                                       for (j in rdata) {
                                          if (rdata[j].action=='redirect') {
                                             var rvalue=rdata[j].value;
                                             processRedirect(rvalue);
                                          }
                                       }
                                    }
                                 });
                              }
                           }
                           wseparator=false;
                        } else {
                           wseparator=true;
                        }
                     }
                  } else {
                     alert("Nieprawidłowy parametr menu.");
                  };
               }
            });
            TREE_OBJ.settings.plugins.contextmenu.items=litems;

            //obsluga zalacznikow
            var anode=TREE_OBJ.get_node(NODE);
            var arel=jQuery(anode).attr("rel");
            if (arel=="attachment") {
               var coid=jQuery("a",anode).attr("coid");
               var roid=jQuery("a",anode).attr("roid");
               i+=1;
               litems['item_'+i.toString()]={
                  label:'Przenieś załącznik',
                  icon:'/icorimg/silk/page_white_copy.png',
                  action: function (NODE,TREE_OBJ,ITEM) {
                     page_params.current_attachment_node=NODE;
                     page_params.current_attachment_coid=coid;
                     page_params.current_attachment_roid=roid;
                  }
               }
            }
            if ((arel=="chapter") && (page_params.current_attachment_coid!="")) {
               //alert("xx : "+page_params.current_attachment_coid);
               var coid=jQuery("a",anode).attr("coid");
               i+=1;
               litems['item_'+i.toString()]={
                  label:'Wstaw załącznik',
                  icon:'/icorimg/silk/page_white_paste.png',
                  separator_before:true,
                  action: function (NODE,TREE_OBJ,ITEM) {
                     jQuery.ajax({
                        url: "icormain.asp?jobtype=workflowmenustructattachmentmovedrag&coid1="+page_params.current_attachment_roid+"&coid2="+coid+"&att="+page_params.current_attachment_coid+"&XMLData=json",
                        async: false,
                        cache: false,
                        type: "GET",
                        dataType: "json",
                        success: function (data, textStatus) {
                           if (textStatus=='success') {
                              if (data.status!='OK') {
                                 alert(data.info);
                              } else {
                                 var aparent=TREE_OBJ.parent(page_params.current_attachment_node);
                                 var anode=NODE;
                                 page_params.current_attachment_node="";
                                 page_params.current_attachment_coid="";
                                 page_params.current_attachment_roid="";
                                 TreeNodeRefresh(anode,TREE_OBJ);
                                 if (aparent) {
                                    //TREE_OBJ.refresh(aparent);
                                    TreeNodeRefresh(aparent,TREE_OBJ);
                                 }
                              }
                           } else {
                              alert("Problem z przeniesieniem załącznika.");
                           };
                        }
                     });
                  }
               }
            }

            i+=1;
            litems['item_'+i.toString()]={
               label:'Odśwież rozdział',
               icon:'/icorimg/silk/arrow_refresh_small.png',
               separator_before:true,
               action: function (NODE,TREE_OBJ,ITEM) {
                  TreeNodeRefresh(NODE,TREE_OBJ);
               }
            }
            i+=1;
            litems['item_'+i.toString()]={
               label:'Odśwież menu',
               icon:'/icorimg/silk/arrow_refresh.png',
               separator_before:true,
               action: function (NODE,TREE_OBJ,ITEM) {
                  window.location.reload();
               }
            }
            amenushow=true;
            if (!amenushow) {
               EV.preventDefault();
               EV.stopPropagation();
            }
            return amenushow;
         }
      }
   });
//   var amenutree=jQuery.tree.reference("#tree_menu");
});
</script>
<%
end if
%>
<%
if 0=1 then
%>
<!-- nodojo
<%
'dim iret
'iret=Process("Introduction")
%>
<tr><td align="left" width="100%"><font face="wingdings" color="red">v</font><a CLASS="BOLD" HREF="icormain.asp?jobtype=lastvisithistory"><span style="color: navy"><font color="navy" face="Arial" onmouseover="this.style.color='blue'" onmouseout="this.style.color='navy'" size="-1"><b>&nbsp;Ostatnio odwiedzane</b></font></span></a></td></tr>
-->
<%
'iret=Process("Contents")
'<!-- ******************** -->
'<script type="text/javascript">
'var node_lastvisited=new WebFXTreeItem("Ostatnio odwiedzane", "", null, "/icormanager/images/wfxtree/items/book_help_navy.png", "/icormanager/images/wfxtree/items/book_open_navy.png");
'tree.add(node_lastvisited);
'function test1() {
'   var anode=new WebFXTreeItem("nowa pozycja", "http://www.google.com", node_lastvisited, "/icormanager/images/wfxtree/items/External_Data_Refresh_Status.png", "/icormanager/images/wfxtree/items/External_Data_Refresh_Status.png");
'   node_lastvisited.add(anode);
'}
'</script>
'<button onclick="javascript:test1();">Tutaj</button>
'<!-- ******************** -->
end if
%>
<!--#include file ="definc_disposeIISI.asp"-->
<SCRIPT LANGUAGE=javascript>
jQuery(function() {
   jQuery.ajaxSetup({
      async:false,
      cache:false
   });
});
</SCRIPT>
<%
if SHOW_TABS=1 then
%>
<SCRIPT LANGUAGE=javascript>
jQuery(function() {
   jQuery('#tabscontainer-1').tabs({selected:0});
});
</SCRIPT>
<%
end if
%>

<SCRIPT LANGUAGE=javascript>
jQuery(function() {
   jQuery("body").append("<div id='mdiv1' class='ui-widget ui-widget-content'></div>");
   var el=jQuery("#mdiv1");
   var abc=el.css("background-color");
   jQuery("html").css("background-color",abc);
   jQuery("body").css("background-color",abc);
   el.remove();
});
</SCRIPT>


<script language="JavaScript" type="text/javascript">
/*
  var onlyOnImages = false;
  var isIE5 = document.all && document.getElementById;
  var isMoz = !isIE5 && document.getElementById;
  function cancelContextMenu(e) {
    if (e && e.stopPropagation)
      e.stopPropagation();
    return false;
  }
  function onContextMenu(e) {
    if (!onlyOnImages
      || (isIE5 && event.srcElement.tagName == "IMG")
      || (IsMoz && e.target.tagName == "IMG")) {
      return cancelContextMenu(e);
    }
  }
  if (document.getElementById) {
    document.oncontextmenu = onContextMenu;
  }
*/
</script>

</body>
</html>