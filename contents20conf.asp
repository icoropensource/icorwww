<%@  codepage="65001" language="VBSCRIPT" %>
<!--#include file ="definc_createIISI.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
   <title>Menu</title>

   <script type="text/javascript" src="/icorlib/jquery/jquery-1.3.2.js"></script>
   <link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/ui-1.7.2/css/<%=Dession("UI_SKIN")%>/jquery-ui-1.7.2.custom.min.css">
   <script type="text/javascript" src="/icorlib/jquery/plugins/ui-1.7.2/js/jquery-ui-1.7.2.custom.js"></script>
   <script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/jquery.tree.js"></script>
   <script type="text/javascript" src="/icorlib/jquery/plugins/jstree/source/plugins/jquery.tree.contextmenu.js"></script>
   <script type="text/javascript" src="/icorlib/json/json2.js"></script>
   <link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/jstree/source/themes/default/style.css">
   <style>
      body, html {
         width: 100%;
         height: 100%;
      }
      body {
         margin: 0px;
         padding: 0px;
         font-size: 12px;
         font-family: 'Trebuchet MS','Arial CE', Arial, sans-serif;
      }
      a img {
         border: none;
      }
      div.menu {
         padding: 2px;
      }
   </style>
   <base target="TEXT">
</head>
<body>

   <div class="demo" id="tree_menu">
   </div>
   <script type="text/javascript">

      function processRedirect(ahref) {
         var apos = ahref.indexOf('!http');
         if (apos >= 0) {
            ahref = ahref.substr(apos + 1);
            window.open(ahref);
            return;
         }
         if (ahref.substr(0, 1) == "!") {
            ahref = ahref.substr(1);
            window.open(ahref);
         } else {
            //alert("3 "+ahref);
            //parent.app.$store.commit('setIframeTextSrc', ahref);
            window.parent.postMessage({
               type: 'setIframeTextSrc',
               ownerTab: '',
               ownerSheet: '',
               href: ahref
            }, '*');
            //window.top.frames("textpane").src=ahref;
         }
      }

      page_params = {
         current_attachment_node: "",
         current_attachment_coid: "",
         current_attachment_roid: ""
      }

      function TreeNodeRefresh(NODE, TREE_OBJ) {
         var w = 0, anode, bnode;
         try {
            var anode = TREE_OBJ.parent(NODE);
            if (anode) {
               TREE_OBJ.refresh(anode);
            } else {
               anode = NODE;
               TREE_OBJ.refresh(anode);
            }
         } catch (e) {
            w = 1;
         };
         if (w == 1) {
            try {
               var bnode = TREE_OBJ.parent(anode);
               if (bnode) {
                  TREE_OBJ.refresh(bnode);
               }
            } catch (e) {
            };
         }
      }

      jQuery(function () {
         jQuery("#tree_menu").tree({
            data: {
               type: "json",
               async: true,
               opts: {
                  async: true,
                  url: "icormain.asp?jobtype=menuxml&OID=-1&group=conf&XMLData=json"
               }
            },
            ui: {
               animation: 0,
               theme_name: 'icor' //'themeroller'
            },
            lang: {
               new_node: "Nowa pozycja",
               loading: "Ładowanie..."
            },
            types: {
               "menustruct": {
                  clickable: true, // can be function
                  renameable: false, // can be function
                  deletable: false, // can be function
                  creatable: false, // can be function
                  draggable: false, // can be function
                  valid_children: ["chapter"] // all, none, array of values can be function
               },
               "chapter": {
                  clickable: true, // can be function
                  renameable: true, // can be function
                  deletable: false, // can be function
                  creatable: false, // can be function
                  draggable: true, // can be function
                  valid_children: ["chapter", "attachment"] // all, none, array of values can be function
               },
               "attachment": {
                  clickable: true, // can be function
                  renameable: false, //true, // can be function
                  deletable: false, // can be function
                  creatable: false, // can be function
                  draggable: false, //true, // can be function
                  valid_children: ["attachment"] // all, none, array of values can be function
               },
               "chaptergroup": {
                  clickable: true, // can be function
                  renameable: true, // can be function
                  deletable: false, // can be function
                  creatable: false, // can be function
                  draggable: false, // can be function
                  valid_children: ["chapter"] // all, none, array of values can be function
               },
               "section": {
                  clickable: true, // can be function
                  renameable: false, // can be function
                  deletable: false, // can be function
                  creatable: false, // can be function
                  draggable: false, // can be function
                  valid_children: ["section"] // all, none, array of values can be function
               }
            },
            plugins: {
               contextmenu: {
                  items: {
                     create: {
                        label: "Nowy rozdział",
                        visible: function (NODE, TREE_OBJ) {
                           return false;
                        }
                     },
                     rename: {
                        label: "Zmień nazwę",
                        visible: function (NODE, TREE_OBJ) {
                           return false;
                        }
                     },
                     remove: {
                        label: "Usuń",
                        visible: function (NODE, TREE_OBJ) {
                           return false;
                        }
                     }
                  }

               }
            },
            callback: {
               onselect: function (NODE, TREE_OBJ) {
                  var anode = TREE_OBJ.get_node(NODE);
                  var aanchor = jQuery("a", anode);
                  var ahref = aanchor.attr("href");
                  if (ahref != "#" && ahref != "") {
                     processRedirect(ahref);
                  }
               },
               beforeopen: function (NODE, TREE_OBJ) {
                  var anode = TREE_OBJ.get_node(NODE);
                  var aanchor = jQuery("a", anode);
                  var aurl = aanchor.attr("submenu");
                  if (aurl == "") {
                     return false;
                  }
                  TREE_OBJ.settings.data.opts.url = aurl;
                  return true
               },
               beforemove: function (NODE, REF_NODE, TYPE, TREE_OBJ) {
                  var anode = TREE_OBJ.get_node(NODE);
                  var arel = jQuery(anode).attr("rel");
                  var coid = jQuery("a", anode).attr("coid");

                  var anoderef = TREE_OBJ.get_node(REF_NODE);
                  var arelref = jQuery(anoderef).attr("rel");
                  var coidref = jQuery("a", anoderef).attr("coid");

                  if ((coid == "") || (coidref == "")) {
                     return false;
                  }
                  //            if (((TYPE=='before') || (TYPE=='after')) && (arelref=='chaptergroup')) {
                  //               return false;
                  //            }
                  if ((arelref == 'menustruct') && ((TYPE == 'before') || (TYPE == 'after'))) {
                     return false;
                  }
                  //            alert(TREE_OBJ.get_text(NODE)+' -> '+TREE_OBJ.get_text(REF_NODE)+' |  '+arel+' | '+arelref+' | '+TYPE);
                  var res = false;
                  if (1 == 1) {
                     jQuery.ajax({
                        url: "icormain.asp?jobtype=workflowmenustructchaptermovedrag&coid1=" + coid + "&coid2=" + coidref + "&type=" + TYPE + "&rel1=" + arel + "&rel2=" + arelref + "&XMLData=json",
                        async: false,
                        cache: false,
                        type: "GET",
                        dataType: "json",
                        success: function (data, textStatus) {
                           if (textStatus == 'success') {
                              if (data.status != 'OK') {
                                 alert(data.info);
                              } else {
                                 res = true;
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
               onrgtclkbefore: function (NODE, TREE_OBJ, EV) {
                  jQuery.tree.plugins.contextmenu.defaults.items = {};
                  TREE_OBJ.settings.plugins.contextmenu.items = {};
                  var amenucontext = jQuery(NODE).find('> a').attr('context');
                  var amenushow = false;
                  var litems = {};
                  var i = 0;
                  jQuery.ajax({
                     url: amenucontext,
                     async: false,
                     cache: false,
                     type: "GET",
                     dataType: "json",
                     success: function (data, textStatus) {
                        if (textStatus == 'success') {
                           var wseparator = false;
                           for (i in data) {
                              if (data[i].text != '-') {
                                 amenushow = true;
                                 litems['item_' + i.toString()] = {
                                    label: data[i].text,
                                    itemaction: data[i].action,
                                    icon: data[i].icon,
                                    separator_before: wseparator,
                                    action: function (NODE, TREE_OBJ, ITEM) {
                                       jQuery.ajax({
                                          url: ITEM.itemaction,
                                          async: false,
                                          cache: false,
                                          type: "GET",
                                          dataType: "json",
                                          success: function (rdata, textStatus) {
                                             for (j in rdata) {
                                                if (rdata[j].action == 'redirect') {
                                                   var rvalue = rdata[j].value;
                                                   processRedirect(rvalue);
                                                }
                                             }
                                          }
                                       });
                                    }
                                 }
                                 wseparator = false;
                              } else {
                                 wseparator = true;
                              }
                           }
                        } else {
                           alert("Nieprawidłowy parametr menu.");
                        };
                     }
                  });
                  TREE_OBJ.settings.plugins.contextmenu.items = litems;

                  //obsluga zalacznikow
                  var anode = TREE_OBJ.get_node(NODE);
                  var arel = jQuery(anode).attr("rel");
                  if (arel == "attachment") {
                     var coid = jQuery("a", anode).attr("coid");
                     var roid = jQuery("a", anode).attr("roid");
                     i += 1;
                     litems['item_' + i.toString()] = {
                        label: 'Przenieś załącznik',
                        icon: '/icorimg/silk/page_white_copy.png',
                        action: function (NODE, TREE_OBJ, ITEM) {
                           page_params.current_attachment_node = NODE;
                           page_params.current_attachment_coid = coid;
                           page_params.current_attachment_roid = roid;
                        }
                     }
                  }
                  if ((arel == "chapter") && (page_params.current_attachment_coid != "")) {
                     //alert("xx : "+page_params.current_attachment_coid);
                     var coid = jQuery("a", anode).attr("coid");
                     i += 1;
                     litems['item_' + i.toString()] = {
                        label: 'Wstaw załącznik',
                        icon: '/icorimg/silk/page_white_paste.png',
                        separator_before: true,
                        action: function (NODE, TREE_OBJ, ITEM) {
                           jQuery.ajax({
                              url: "icormain.asp?jobtype=workflowmenustructattachmentmovedrag&coid1=" + page_params.current_attachment_roid + "&coid2=" + coid + "&att=" + page_params.current_attachment_coid + "&XMLData=json",
                              async: false,
                              cache: false,
                              type: "GET",
                              dataType: "json",
                              success: function (data, textStatus) {
                                 if (textStatus == 'success') {
                                    if (data.status != 'OK') {
                                       alert(data.info);
                                    } else {
                                       var aparent = TREE_OBJ.parent(page_params.current_attachment_node);
                                       var anode = NODE;
                                       page_params.current_attachment_node = "";
                                       page_params.current_attachment_coid = "";
                                       page_params.current_attachment_roid = "";
                                       TreeNodeRefresh(anode, TREE_OBJ);
                                       if (aparent) {
                                          //TREE_OBJ.refresh(aparent);
                                          TreeNodeRefresh(aparent, TREE_OBJ);
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

                  i += 1;
                  litems['item_' + i.toString()] = {
                     label: 'Odśwież rozdział',
                     icon: '/icorimg/silk/arrow_refresh_small.png',
                     separator_before: true,
                     action: function (NODE, TREE_OBJ, ITEM) {
                        TreeNodeRefresh(NODE, TREE_OBJ);
                     }
                  }
                  i += 1;
                  litems['item_' + i.toString()] = {
                     label: 'Odśwież menu',
                     icon: '/icorimg/silk/arrow_refresh.png',
                     separator_before: true,
                     action: function (NODE, TREE_OBJ, ITEM) {
                        window.location.reload();
                     }
                  }
                  amenushow = true;
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

   <!--#include file ="definc_disposeIISI.asp"-->
   <script language="javascript">
      jQuery(function () {
         jQuery.ajaxSetup({
            async: true,
            cache: false
         });
      });
   </script>
   <script language="javascript">
      jQuery(function () {
         jQuery("html").css("background-color", '#efefff');
         jQuery("body").css("background-color", '#efefff');
      });
   </script>

</body>
</html>
