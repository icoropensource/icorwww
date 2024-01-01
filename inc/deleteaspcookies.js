 //Clear any session cookies
 (function(){
     var i;
     var cookiesArr;
     var cItem;

     cookiesArr = document.cookie.split("; ");
     for (i = 0; i < cookiesArr.length; i++) {
         cItem = cookiesArr[i].split("=");
         if (cItem.length > 0 && cItem[0].indexOf("ASPSESSIONID") == 0) {
             deleteCookie(cItem[0]);
         }
     }

     function deleteCookie(name) {
         var expDate = new Date();
         expDate.setTime(expDate.getTime() - 86400000); //-1 day
         var value = "; expires=" + expDate.toGMTString() + ";path=/";
         document.cookie = name + "=" + value;
     }
 })();
