/*----------------------------------------------------------------------------\
|                            Sortable Table 1.03                              |
\----------------------------------------------------------------------------*/


function SortableTable(oTable, oSortTypes) {

   this.element = oTable;
   this.tHead = oTable.tHead;
   this.tBody = oTable.tBodies[0];
   this.document = oTable.ownerDocument || oTable.document;
   
   this.sortColumn = null;
   this.descending = null;
   
   var oThis = this;
   this._headerOnclick = function (e) {
      oThis.headerOnclick(e);
   };
   
   
   // only IE needs this
   var win = this.document.defaultView || this.document.parentWindow;
   this._onunload = function () {
      oThis.destroy();
   };
   if (win && typeof win.attachEvent != "undefined") {
      win.attachEvent("onunload", this._onunload);
   }
   
   this.initHeader(oSortTypes || []);
}

SortableTable.gecko = navigator.product == "Gecko";
SortableTable.msie = /msie/i.test(navigator.userAgent);
// Mozilla is faster when doing the DOM manipulations on
// an orphaned element. MSIE is not
SortableTable.removeBeforeSort = SortableTable.gecko;

SortableTable.prototype.onsort = function () {};

// adds arrow containers and events
// also binds sort type to the header cells so that reordering columns does
// not break the sort types
SortableTable.prototype.initHeader = function (oSortTypes) {
   var cells = this.tHead.rows[0].cells;
   var l = cells.length;
   var img, c;
   for (var i = 0; i < l; i++) {
      c = cells[i];
      img = this.document.createElement("IMG");
      img.src = "/icormanager/images/sorttable_blank.png";
      c.appendChild(img);
      if (oSortTypes[i] != null) {
         c._sortType = oSortTypes[i];
      }
      if (typeof c.addEventListener != "undefined")
         c.addEventListener("click", this._headerOnclick, false);
      else if (typeof c.attachEvent != "undefined")      
         c.attachEvent("onclick", this._headerOnclick);
   }
   this.updateHeaderArrows();
};

// remove arrows and events
SortableTable.prototype.uninitHeader = function () {
   var cells = this.tHead.rows[0].cells;
   var l = cells.length;
   var c;
   for (var i = 0; i < l; i++) {
      c = cells[i];
      c.removeChild(c.lastChild);
      if (typeof c.removeEventListener != "undefined")
         c.removeEventListener("click", this._headerOnclick, false);
      else if (typeof c.detachEvent != "undefined")
         c.detachEvent("onclick", this._headerOnclick);
   }
};

SortableTable.prototype.updateHeaderArrows = function () {
   var cells = this.tHead.rows[0].cells;
   var l = cells.length;
   var img;
   for (var i = 0; i < l; i++) {
      img = cells[i].lastChild;
      if (i == this.sortColumn)
         img.className = "sort-arrow " + (this.descending ? "descending" : "ascending");
      else
         img.className = "sort-arrow";       
   }
};

SortableTable.prototype.headerOnclick = function (e) {
   // find TD element
   var el = e.target || e.srcElement;
   while (el.tagName != "TD")
      el = el.parentNode;
   
   this.sort(el.cellIndex);   
};

SortableTable.prototype.getSortType = function (nColumn) {
   var cell = this.tHead.rows[0].cells[nColumn];
   var val = cell._sortType;
   if (val != "")
      return val;
   return "String";
};

// only nColumn is required
// if bDescending is left out the old value is taken into account
// if sSortType is left out the sort type is found from the sortTypes array

SortableTable.prototype.sort = function (nColumn, bDescending, sSortType) {
   if (sSortType == null)
      sSortType = this.getSortType(nColumn);

   // exit if None   
   if (sSortType == "None")
      return;
   
   if (bDescending == null) {
      if (this.sortColumn != nColumn)
         this.descending = true;
      else
         this.descending = !this.descending;
   }  
   
   this.sortColumn = nColumn;
   
   if (typeof this.onbeforesort == "function")
      this.onbeforesort();
   
   var f = this.getSortFunction(sSortType, nColumn);
   var a = this.getCache(sSortType, nColumn);
   var tBody = this.tBody;
   
   a.sort(f);
   
   if (this.descending)
      a.reverse();
   
   if (SortableTable.removeBeforeSort) {
      // remove from doc
      var nextSibling = tBody.nextSibling;
      var p = tBody.parentNode;
      p.removeChild(tBody);
   }
   
   // insert in the new order
   var l = a.length;
   for (var i = 0; i < l; i++)
      tBody.appendChild(a[i].element);
   
   if (SortableTable.removeBeforeSort) {  
      // insert into doc
      p.insertBefore(tBody, nextSibling);
   }
   
   this.updateHeaderArrows();
   
   this.destroyCache(a);
   
   if (typeof this.onsort == "function")
      this.onsort();
};

SortableTable.prototype.asyncSort = function (nColumn, bDescending, sSortType) {
   var oThis = this;
   this._asyncsort = function () {
      oThis.sort(nColumn, bDescending, sSortType);
   };
   window.setTimeout(this._asyncsort, 1); 
};

SortableTable.prototype.getCache = function (sType, nColumn) {
   var rows = this.tBody.rows;
   var l = rows.length;
   var a = new Array(l);
   var r;
   for (var i = 0; i < l; i++) {
      r = rows[i];
      a[i] = {
         value:      this.getRowValue(r, sType, nColumn),
         element: r
      };
   };
   return a;
};

SortableTable.prototype.destroyCache = function (oArray) {
   var l = oArray.length;
   for (var i = 0; i < l; i++) {
      oArray[i].value = null;
      oArray[i].element = null;
      oArray[i] = null;
   }
}

SortableTable.prototype.getRowValue = function (oRow, sType, nColumn) {
   var s;
   var c = oRow.cells[nColumn];
   if (typeof c.innerText != "undefined")
      s = c.innerText;
   else
      s = SortableTable.getInnerText(c);
   return this.getValueFromString(s, sType);
};

SortableTable.getInnerText = function (oNode) {
   var s = ""; 
   var cs = oNode.childNodes;
   var l = cs.length;
   for (var i = 0; i < l; i++) {
      switch (cs[i].nodeType) {
         case 1: //ELEMENT_NODE
            s += SortableTable.getInnerText(cs[i]);
            break;
         case 3:  //TEXT_NODE
            s += cs[i].nodeValue;
            break;
      }
   }
   return s;
}

SortableTable.prototype.getValueFromString = function (sText, sType) {
   var apatt=/\<[^\>]+>/g;
   sText=sText.replace(apatt,"");
   switch (sType) {
      case "Number":
         return Number(sText);
      case "CaseInsensitiveString":
         return sText.toUpperCase();
      case "Date":
         var d=new Date(0);
         apatt=/(\d+)[\/\-](\d+)[\/\-](\d+)/;
         var a=apatt.exec(sText);
         if (a) {
            var y=parseInt(RegExp.$1);
            if (y>31) {
               d.setFullYear(RegExp.$1);
               d.setDate(RegExp.$3);
               d.setMonth(RegExp.$2-1);
            } else {
               d.setFullYear(RegExp.$3);
               d.setDate(RegExp.$1);
               d.setMonth(RegExp.$2-1);
            }
         }
         apatt=/(\d+)\:(\d+)\:(\d+)/;
         var a=apatt.exec(sText);
         if (a) {
            d.setHours(parseInt(RegExp.$1),parseInt(RegExp.$2),parseInt(RegExp.$3));
         }
         return d.valueOf();
   }
   return sText;
};

SortableTable.prototype.getSortFunction = function (sType, nColumn) {
   return function compare(n1, n2) {
      if (n1.value < n2.value)
         return -1;
      if (n2.value < n1.value)
         return 1;
      return 0;
   };
};

SortableTable.prototype.destroy = function () {
   this.uninitHeader();
   var win = this.document.parentWindow;
   if (win && typeof win.detachEvent != "undefined") {   // only IE needs this
      win.detachEvent("onunload", this._onunload);
   }  
   this._onunload = null;
   this.element = null;
   this.tHead = null;
   this.tBody = null;
   this.document = null;
   this._headerOnclick = null;
   this.sortTypes = null;
   this._asyncsort = null;
   this.onsort = null;
};
