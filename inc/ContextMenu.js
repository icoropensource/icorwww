ContextMenu.intializeContextMenu=function()
{
   document.body.insertAdjacentHTML("BeforeEnd", '<iframe scrolling="no" class="WebFX-ContextMenu" marginwidth="0" marginheight="0" frameborder="0" style="position:absolute;display:none;z-index:50000000;" id="WebFX_PopUp"></iframe>');
   WebFX_PopUp    = self.frames["WebFX_PopUp"]
   WebFX_PopUpcss = document.getElementById("WebFX_PopUp")
   document.body.attachEvent("onmousedown",function(){WebFX_PopUpcss.style.display="none"})
   WebFX_PopUpcss.onfocus  = function(){WebFX_PopUpcss.style.display="inline"};
   WebFX_PopUpcss.onblur  = function(){WebFX_PopUpcss.style.display="none"};
   self.attachEvent("onblur",function(){WebFX_PopUpcss.style.display="none"})
}


function ContextSeperator(){}

function ContextMenu(){}

ContextMenu.showPopup=function(x,y)
{
   WebFX_PopUpcss.style.display = "block"
}

ContextMenu.hidePopup=function()
{
   WebFX_PopUpcss.style.display = "none"
}

ContextMenu.display=function(popupoptions,eobj)
{
   var x,y;
   if (!eobj) {
      eobj = window.event;
      x    = eobj.clientX;
      y    = eobj.clientY;
   } else {
      x    = eobj.clientX;
      y    = eobj.clientY;
   }
   
   /*
   not really sure why I had to pass window here
   it appears that an iframe inside a frames page
   will think that its parent is the frameset as
   opposed to the page it was created in...
   */
   ContextMenu.populatePopup(popupoptions,window)
   
   ContextMenu.showPopup(x,y)
   ContextMenu.fixSize()
   ContextMenu.fixPos(x,y)
   eobj.cancelBubble = true;
   eobj.returnValue  = false;
}

//TODO
 ContextMenu.getScrollTop=function()
 {
   return document.documentElement.scrollTop;
//   return document.body.scrollTop;
   //window.pageXOffset and window.pageYOffset for moz
 }
 
 ContextMenu.getScrollLeft=function()
 {
   return document.documentElement.scrollLeft;
//   return document.body.scrollLeft;
 }
 


ContextMenu.fixPos=function(x,y)
{
   var docheight,docwidth,dh,dw;
   docheight = document.body.clientHeight;
   docwidth  = document.body.clientWidth;
   dh = (WebFX_PopUpcss.offsetHeight+y) - docheight;
   dw = (WebFX_PopUpcss.offsetWidth+x)  - docwidth;
   //alert("x: "+ContextMenu.getScrollLeft()+" : "+ContextMenu.getScrollTop());
   var xdw=x;
   if (dw>0) {
      xdw=x-dw;
      if (xdw<0) {
        xdw=0;
      };
   };
   WebFX_PopUpcss.style.left=xdw+ContextMenu.getScrollLeft()+"px";
   var ydh=y;
   if (dh>0) {
      ydh=y-dh;
      if (ydh<0) {
        ydh=0;
      };
   };
   WebFX_PopUpcss.style.top=ydh+ContextMenu.getScrollTop()+"px";
}

ContextMenu.fixPosOLD=function(x,y)
{
   var docheight,docwidth,dh,dw;
   docheight = document.body.clientHeight;
   docwidth  = document.body.clientWidth;
   dh = (WebFX_PopUpcss.offsetHeight+y) - docheight;
   dw = (WebFX_PopUpcss.offsetWidth+x)  - docwidth;
//   alert("x: "+ContextMenu.getScrollLeft()+" : "+ContextMenu.getScrollTop());
   if(dw>0)
   {
      WebFX_PopUpcss.style.left = (x - dw) + ContextMenu.getScrollLeft() + "px";    
   }
   else
   {
      WebFX_PopUpcss.style.left = x + ContextMenu.getScrollLeft();
   }
   if(dh>0)
   {
      WebFX_PopUpcss.style.top = (y - dh) + ContextMenu.getScrollTop() + "px"
   }
   else
   {
      WebFX_PopUpcss.style.top  = y + ContextMenu.getScrollTop();
   }
}

ContextMenu.fixSize=function()
{
   var body,h,w;
   WebFX_PopUpcss.style.width = "10px";
   WebFX_PopUpcss.style.height = "100000px";
   body = WebFX_PopUp.document.body; 
   // check offsetHeight twice... fixes a bug where scrollHeight is not valid because the visual state is undefined
   var dummy = WebFX_PopUpcss.offsetHeight + " dummy";
   h = body.scrollHeight + WebFX_PopUpcss.offsetHeight - body.clientHeight;
   w = body.scrollWidth + WebFX_PopUpcss.offsetWidth - body.clientWidth;
   WebFX_PopUpcss.style.height = h + "px";
   WebFX_PopUpcss.style.width = w + "px";
   //use document.height for moz
}

ContextMenu.populatePopup=function(arr,win)
{
   var alen,i,tmpobj,doc,height,htmstr;
   alen = arr.length;
   doc  = WebFX_PopUp.document;
   doc.body.innerHTML  = ""
   if (doc.getElementsByTagName("LINK").length == 0) {
      doc.open();
      doc.write('<html><head><link rel="StyleSheet" type="text/css" href="/icormanager/inc/WebFX-ContextMenu.css"><meta HTTP-EQUIV="Content-Type" content="text/html; charset=windows-1250"><meta http-equiv="Content-Language" content="pl"></head><body></body></html>');
      doc.close();
   }
   for(i=0;i<alen;i++)
   {
//      alert('bb '+arr[i]);
      if(arr[i].constructor==ContextItem)
      {
         tmpobj=doc.createElement("DIV");
         tmpobj.noWrap = true;
         tmpobj.className = "WebFX-ContextMenu-Item";

         if(arr[i].disabled)
         {
            htmstr  = '<span class="WebFX-ContextMenu-DisabledContainer"><span class="WebFX-ContextMenu-DisabledContainer">';
            htmstr += arr[i].text+'</span></span>';
            tmpobj.innerHTML = htmstr;
            tmpobj.className = "WebFX-ContextMenu-Disabled";
            tmpobj.onmouseover = function(){this.className="WebFX-ContextMenu-Disabled-Over"};
            tmpobj.onmouseout  = function(){this.className="WebFX-ContextMenu-Disabled"};
         }
         else
         {
            tmpobj.innerHTML = arr[i].text;
            tmpobj.onclick = (function (f,p,win)
                        {
                               return function () {
                                             win.WebFX_PopUpcss.style.display='none';
                                             if (typeof(f)=="function") {
                                                f(p);
                                             }
                               };
                        })(arr[i].action,arr[i].param,win);
               
            tmpobj.onmouseover = function(){this.className="WebFX-ContextMenu-Over"};
            tmpobj.onmouseout  = function(){this.className="WebFX-ContextMenu-Item"};
         }
         doc.body.appendChild(tmpobj);
      }
      else
      {
         doc.body.appendChild(doc.createElement("DIV")).className = "WebFX-ContextMenu-Separator";
      }
   }
   doc.body.className  = "WebFX-ContextMenu-Body" ;
   doc.body.onselectstart = function(){return false;}
}

function ContextItem(str,fnc,disabled,param)
{
   this.text     = str;
   this.action   = fnc; 
   this.disabled = disabled || false;
   this.param    = param || false;
}
