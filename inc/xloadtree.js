/*----------------------------------------------------------------------------\
|                               XLoadTree 1.1                                 |
\----------------------------------------------------------------------------*/

webFXTreeConfig.loadingText = "£adowanie danych...";
webFXTreeConfig.loadErrorTextTemplate = "Wyst¹pi³ b³¹d \"%1%\"";
webFXTreeConfig.emptyErrorTextTemplate = "\"%1%\" nie zawiera ¿adnych podpozycji";

/*
 * WebFXLoadTree class
 */

function WebFXLoadTree(sText, sXmlSrc, sAction, sBehavior, sIcon, sOpenIcon) {
   // call super
   this.WebFXTree = WebFXTree;
   this.WebFXTree(sText, sAction, sBehavior, sIcon, sOpenIcon);
   
   // setup default property values
   this.src = sXmlSrc;
   this.loading = false;
   this.loaded = false;
   this.errorText = "";
   
   // check start state and load if open
   if (this.open)
      _startLoadXmlTree(this.src, this);
   else {
      // and create loading item if not
      this._loadingItem = new WebFXTreeItem(webFXTreeConfig.loadingText);
      this.add(this._loadingItem);
   }
}

WebFXLoadTree.prototype = new WebFXTree;

// override the expand method to load the xml file
WebFXLoadTree.prototype._webfxtree_expand = WebFXTree.prototype.expand;
WebFXLoadTree.prototype.expand = function() {
   if (!this.loaded && !this.loading) {
      // load
      _startLoadXmlTree(this.src, this);
   }
   this._webfxtree_expand();
};

/*
 * WebFXLoadTreeItem class
 */

function WebFXLoadTreeItem(sText, sXmlSrc, sAction, eParent, sIcon, sOpenIcon, sContext) {
   // call super
   this.WebFXTreeItem = WebFXTreeItem;
   this.WebFXTreeItem(sText, sAction, eParent, sIcon, sOpenIcon, sContext);

   // setup default property values
   this.src = sXmlSrc;
   this.loading = false;
   this.loaded = false;
   this.errorText = "";
   
   // check start state and load if open
   if (this.open){
      _startLoadXmlTree(this.src, this);}
   else {
      // and create loading item if not
      this._loadingItem = new WebFXTreeItem(webFXTreeConfig.loadingText);
      this.add(this._loadingItem);
   }
}

WebFXLoadTreeItem.prototype = new WebFXTreeItem;

// override the expand method to load the xml file
WebFXLoadTreeItem.prototype._webfxtreeitem_expand = WebFXTreeItem.prototype.expand;
WebFXLoadTreeItem.prototype.expand = function() {
   if (!this.loaded && !this.loading) {
      // load
      _startLoadXmlTree(this.src, this);
   }
   this._webfxtreeitem_expand();
   if (this.refresh){
      this.loading = false;
      this.reload();
      _startLoadXmlTree(this.src, this);    
      this.refresh=false;
   }
   
};

// reloads the src file if already loaded
WebFXLoadTree.prototype.reload = 
WebFXLoadTreeItem.prototype.reload = function () {
   // if loading do nothing
   this.loadingCnt=0;
   if (this.loaded) {
      var open = this.open;
      // remove
      while (this.childNodes.length > 0)
         this.childNodes[this.childNodes.length - 1].remove();
      
      this.loaded = false;
      
      this._loadingItem = new WebFXTreeItem(webFXTreeConfig.loadingText);
      this.add(this._loadingItem);
      
      if (open)
         this.expand();         
   }
   else if (this.open && !this.loading)
      _startLoadXmlTree(this.src, this);
};

/*
 * Helper functions
 */

// creates the xmlhttp object and starts the load of the xml document
function _startLoadXmlTree(sSrc, jsNode) {
   if (jsNode.loading || jsNode.loaded)
      return;
   jsNode.loading = true;
   var xmlHttp = XmlHttp.create();
   xmlHttp.open("GET", sSrc, true); // async
   xmlHttp.onreadystatechange = function () {
      if (xmlHttp.readyState == 4) {
         _xmlFileLoaded(xmlHttp.responseXML, jsNode, xmlHttp.responseText);
      }
   };

   
   
   
   // call in new thread to allow ui to update
   window.setTimeout(function () {
      xmlHttp.send(null);
/*      var tmStart = new Date();
      var tmCurr = new Date();
      tmStart.getTime();
      while (xmlHttp.readyState!=4) {
         tmCurr.getTime();
         if ((tmCurr-tmStart)>1000) {
            xmlHttp.abort();
            alert("babol!");
            break;
         }
      } */
   }, 10);
}

function getAsPLText(atext) {
   if (atext==null) {
      atext="-***-";
   }
   while (atext.search("&#185;")>=0)
      atext=atext.replace("&#185;","¹");
   while (atext.search("&#230;")>=0)
      atext=atext.replace("&#230;","æ");
   while (atext.search("&#234;")>=0)
      atext=atext.replace("&#234;","ê");
   while (atext.search("&#179;")>=0)
      atext=atext.replace("&#179;","³");
   while (atext.search("&#241;")>=0)
      atext=atext.replace("&#241;","ñ");
   while (atext.search("&#243;")>=0)
      atext=atext.replace("&#243;","ó");
   while (atext.search("&#156;")>=0)
      atext=atext.replace("&#156;","œ");
   while (atext.search("&#159;")>=0)
      atext=atext.replace("&#159;","Ÿ");
   while (atext.search("&#191;")>=0)
      atext=atext.replace("&#191;","¿");
   while (atext.search("&#165;")>=0)
      atext=atext.replace("&#165;","¥");
   while (atext.search("&#198;")>=0)
      atext=atext.replace("&#198;","Æ");
   while (atext.search("&#202;")>=0)
      atext=atext.replace("&#202;","Ê");
   while (atext.search("&#163;")>=0)
      atext=atext.replace("&#163;","£");
   while (atext.search("&#209;")>=0)
      atext=atext.replace("&#209;","Ñ");
   while (atext.search("&#211;")>=0)
      atext=atext.replace("&#211;","Ó");
   while (atext.search("&#140;")>=0)
      atext=atext.replace("&#140;","Œ");
   while (atext.search("&#143;")>=0)
      atext=atext.replace("&#143;","");
   while (atext.search("&#175;")>=0)
      atext=atext.replace("&#175;","¯");
   return atext;
}

// Converts an xml tree to a js tree. See article about xml tree format
function _xmlTreeToJsTree(oNode) {
   // retreive attributes
   var text = getAsPLText(oNode.getAttribute("text"));
   var action = oNode.getAttribute("action");
   var parent = null;
   var icon = oNode.getAttribute("icon");
   var openIcon = oNode.getAttribute("openIcon");
   var src = oNode.getAttribute("src");
   var context = oNode.getAttribute("context");
   // create jsNode
   var jsNode;
   if (src != null && src != "") {
      jsNode = new WebFXLoadTreeItem(text, src, action, parent, icon, openIcon, context);
   } else {
      jsNode = new WebFXTreeItem(text, action, parent, icon, openIcon, context);
   }
//$$<
   var allowmove = oNode.getAttribute("allowmove");
   jsNode.allowmove=allowmove;
   var coid = oNode.getAttribute("coid");
   jsNode.coid=coid;
//$$>
   // go through childNOdes
   var cs = oNode.childNodes;
   var l = cs.length;
   for (var i = 0; i < l; i++) {
      if (cs[i].tagName == "tree")
         jsNode.add( _xmlTreeToJsTree(cs[i]), true );
   }
   return jsNode;
}

// Inserts an xml document as a subtree to the provided node
function _xmlFileLoaded(oXmlDoc, jsParentNode, atext) {
   if (jsParentNode.loaded)
      return;
   var bIndent = false;
   var bAnyChildren = false;
   jsParentNode.loaded = true;
   jsParentNode.loading = false;
   // check that the load of the xml file went well
   if( oXmlDoc == null || oXmlDoc.documentElement == null) {
      if( oXmlDoc == null) {
         alert("oXmlDoc == null");
      }
      if( oXmlDoc.documentElement == null) {
         jsParentNode.loaded = false;
         if (jsParentNode.loadingCnt<3) {
            jsParentNode.loadingCnt=jsParentNode.loadingCnt+1;
            _startLoadXmlTree(jsParentNode.src, jsParentNode);
         }
         jsParentNode.collapse();
         return
      }
      jsParentNode.errorText = parseTemplateString(webFXTreeConfig.loadErrorTextTemplate,jsParentNode.src);
   }
   else {
      // there is one extra level of tree elements
      var root = oXmlDoc.documentElement;

      // loop through all tree children
      var cs = root.childNodes;
      var l = cs.length;
      for (var i = 0; i < l; i++) {
         if (cs[i].tagName == "tree") {
            bAnyChildren = true;
            bIndent = true;
            jsParentNode.add( _xmlTreeToJsTree(cs[i]), true);
         }
      }

      // if no children we got an error
      if (!bAnyChildren)
         jsParentNode.errorText = parseTemplateString(webFXTreeConfig.emptyErrorTextTemplate,
                              jsParentNode.src);
   }
   
   // remove dummy
   if (jsParentNode._loadingItem != null) {
      jsParentNode._loadingItem.remove();
      bIndent = true;
   }
   
   if (bIndent) {
      // indent now that all items are added
      jsParentNode.indent();
   }
   jsParentNode.loadingCnt=0;
   
   // show error in status bar
   if (jsParentNode.errorText != "")
      window.status = jsParentNode.errorText;
}

// parses a string and replaces %n% with argument nr n
function parseTemplateString(sTemplate) {
   var args = arguments;
   var s = sTemplate;
   
   s = s.replace(/\%\%/g, "%");
   
   for (var i = 1; i < args.length; i++)
      s = s.replace( new RegExp("\%" + i + "\%", "g"), args[i] )
   
   return s;
}
