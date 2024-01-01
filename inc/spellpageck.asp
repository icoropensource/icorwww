<%@ CODEPAGE=65001 LANGUAGE="VBSCRIPT" %>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta name="pragma" content="no-cache">
<META HTTP-EQUIV="expires" CONTENT="Mon, 1 Jan 2001 01:01:01 GMT"> 
<TITLE>Sprawdzanie pisowni</TITLE>
<link rel=STYLESHEET type="text/css" href="/icormanager/icor.css" title="SOI">

<script language='javascript'>
var drPopup = window.createPopup();
var lastelement=-1;
var changeall=0;
var alastoriginal;
var aimgelementscount=0;
var ahtml;
var asourcetextelement;
var alanguage;
var langname;
function showPopupDR(asrc,aelement) {
   lastelement=aelement;
   drPopup.document.body.innerHTML = asrc.innerHTML;
   drPopup.show(0, 18, 200, 155, aelement);
}
function processPopup(avalue) {
   if (drPopup.isOpen)
      drPopup.hide();
   var aoriginal=lastelement.getAttribute("mOriginal");
   alastoriginal=lastelement.getAttribute("mLastOriginal");
   
   aframedocument=document.all("myspellframe").contentWindow.document;
   adocumentbody=aframedocument.body;
   atextrange=adocumentbody.createTextRange();
   alen=parseInt(lastelement.getAttribute("mLen"));
   apos=parseInt(lastelement.getAttribute("mPos"));
   if (changeall==1){
      if (alastoriginal != '')
         while (atextrange.findText(alastoriginal,1,6) ){
               atextrange.pasteHTML(avalue);
         }
       for (i=0;i<aimgelementscount;i++) {
          try{
             aelement=aframedocument.all("myimg_"+i);
             if (aelement.getAttribute('mLastOriginal')==alastoriginal){
                aelement.setAttribute('mLastOriginal',avalue);
                aelement.setAttribute('mLen',avalue.length.toString());
                if (aoriginal==avalue){
                    aelement.setAttribute('src','../images/SpellLetterRed.png');
                } else {
                    aelement.setAttribute('src','../images/SpellLetterGreen.png');           
                }    
             }   
          }
          catch(e){}
       }    
   }
   else{
      atextrange.moveToElementText(lastelement);
      atextrange.moveStart("character",-alen);
      atextrange.moveEnd("character",-1);
      atextrange.text=avalue;
      if (aoriginal==avalue){
          lastelement.setAttribute('src','../images/SpellLetterRed.png');
      } else{
          lastelement.setAttribute('src','../images/SpellLetterGreen.png');           
      }    
      lastelement.setAttribute("mLen",avalue.length.toString());
      lastelement.setAttribute("mLastOriginal",avalue);   
   }
   document.all("lastclickedword").value=avalue
}
function processMTCommand(aelement) {
   var aoffset=aelement.getAttribute("mOffset");
   asrc=document.all("div_"+aoffset);
   lastelement.border=0;
   aelement.border=1;
   document.all("lastclickedword").value=aelement.getAttribute("mLastOriginal");
   document.all('idnrbledu').innerHTML=parseInt(aelement.getAttribute('nr'))+1;   
   changeall=0;
   showPopupDR(asrc,aelement);
}
function findMistake(awhere){
  aframedocument=document.all('myspellframe').contentWindow.document;
  ascroll=0;
  switch(awhere){
    case 0:
       for (i=0;i<aimgelementscount;i++) {
          try{
             aelement=aframedocument.all("myimg_"+i);
             if ((aelement.getAttribute('mLastOriginal')==aelement.getAttribute('mOriginal')) && (aelement.getAttribute('add')=='0')){
                ascroll=1;
                break;
             }
          }
          catch(e){}
       }    
       break;
    case 1:
       acurrentimagenr=parseInt(lastelement.getAttribute('nr'));
       for (i=acurrentimagenr-1;i>=0;i--) {
          try{
             aelement=aframedocument.all("myimg_"+i);
             if ((aelement.getAttribute('mLastOriginal')==aelement.getAttribute('mOriginal')) && (aelement.getAttribute('add')=='0')){
                ascroll=1;
                break;
             }
          }
          catch(e){}
       }    
       break;
    case 2:
       acurrentimagenr=parseInt(lastelement.getAttribute('nr'));
       for (i=acurrentimagenr+1;i<aimgelementscount;i++) {
          try{
             aelement=aframedocument.all("myimg_"+i);
             if ((aelement.getAttribute('mLastOriginal')==aelement.getAttribute('mOriginal')) && (aelement.getAttribute('add')=='0')){
                ascroll=1;
                break;
             }
          }
          catch(e){}
       }    
       break;
    case 3:
       for (i=aimgelementscount-1;i>0;i--) {
          try{
             aelement=aframedocument.all("myimg_"+i);
             if ((aelement.getAttribute('mLastOriginal')==aelement.getAttribute('mOriginal')) && (aelement.getAttribute('add')=='0')){
                ascroll=1;
                break;
             }
          }
          catch(e){}
       }    
  }
  if (ascroll==1){
     aelement.scrollIntoView();
     document.all('idnrbledu').innerHTML=parseInt(aelement.getAttribute('nr'))+1;
     document.all("lastclickedword").value=aelement.getAttribute('mLastOriginal');
     lastelement.border=0;
     aelement.border=1;
     lastelement=aelement;
  }
}
function doClearImages() {
   aframedocument=document.all("myspellframe").contentWindow.document;
   for (i=0;i<aimgelementscount;i++) {
      try{
         aelement=aframedocument.all("myimgspan_"+i);
         aelement.removeNode(true);
      }
      catch(e){}      
   }
   aimgelementscount=0;
}

function replacehtml(){
  document.all('divmessage').style.top=(window.document.body.clientHeight-40)/2;
  document.all('divmessage').style.left=(window.document.body.clientWidth-300)/2;  
  awindowparams=window.dialogArguments;  
  //asourcetextelement=awindowparams.element;
  asourcetextelement=awindowparams.html //$$SK  
  alanguage=awindowparams.lang;
  document.all('lastclickedword').value='';
  aimgelementscount=0;
  aframedocument=document.all("myspellframe").contentWindow.document;  
  adocumentbody=aframedocument.body;    
  adocumentbody.innerHTML=asourcetextelement;
  
  if (adocumentbody.innerHTML==''){
     window.returnValue='';
     window.close();
     return;
  }  
  aretdiv=new Array();
  aretdivpos=0;
  atextrange=adocumentbody.createTextRange();
  
  switch(alanguage){
    case 'pl':
       langname='polski';
       break;
    case 'en':
       langname='angielski';    
       break;   
    case 'de':
       langname='niemiecki';    
  }  
  document.all('idlang').innerHTML='<img style="border-style: solid; border-width: 1" src="/icormanager/images/flags/lang_'+alanguage+'.png" alt="'+langname+'">';  
  
  var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  xmlHttp.open('POST', 'SpellCheckXML.asp', true);
  xmlHttp.onreadystatechange = function () {
     if (xmlHttp.readyState == 4) {
        replaceHTMLStep2(xmlHttp,atextrange);
      }
  };

  window.setTimeout(function (){ xmlHttp.send('<?xml version="1.0" encoding="utf-8" ?> <DANE mode="check" language="'+alanguage+'"><![CDATA[' + atextrange.text + ']]></DANE>')});  
}

function replaceHTMLStep2(xmlHttp,atextrange) {
  if (xmlHttp.responseText=='NOT AVAILABLE'){
     alert('Funkcja nieaktywna!');
     window.returnValue='';
     window.close();
     return;
  }
  if (xmlHttp.responseText=='ERROR'){
     alert('Problem z serwerem. Powiadom administratora!');
  } else{
  var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
  xmlDoc.async=false;
  xmlDoc.loadXML(xmlHttp.responseText);
  lines=xmlDoc.getElementsByTagName("L");
  document.all('idliczbaslow').innerHTML=xmlDoc.getElementsByTagName("STATS").item(0).getAttribute("wordcount");
  document.all('idczas').innerHTML=xmlDoc.getElementsByTagName("STATS").item(0).getAttribute("time");
  nlines=lines.length;
  nmistakes=0;
  linestart=0;

  for(i=0; i<nlines; i++){
     line= lines.item(i);
     linelength=line.getAttribute("l");
     misstakes=line.childNodes;
     for(j=0; j<misstakes.length; j++){
        nmistakes=nmistakes+1;
        misstake= misstakes.item(j);
        wordoffset=misstake.getAttribute("w");
        original=misstake.getAttribute("o");
        offset=misstake.getAttribute("f");
        sugestions=misstake.childNodes;
        nsugestions=sugestions.length;
        aretdiv[aretdivpos++]="<div id=div_"+wordoffset+" name=div_"+wordoffset+" STYLE='display:none;'>";
        aretdiv[aretdivpos++]="<div style='position:absolute; top:0; left:0;width:200;height:45;background:wheat;'><center><br style='line-height:30%;'><button style='width:100pt;height:14pt;FONT-FAMILY: arial, verdana;FONT-SIZE: 8pt;FONT-STYLE: normal;FONT-WEIGHT: normal;' onclick='javascript:parent.doAdd(\""+original.replace(/\'/gi,"&#39;")+"\")'>Dodaj do s³ownika</button><br><font style='FONT-FAMILY: arial, verdana;FONT-SIZE: 8pt;FONT-STYLE: normal;FONT-WEIGHT: normal;'>Zamieñ wszystkie</font><input type=checkbox id=checkreplaceall name=checkreplaceall onclick='javascript:parent.changeall=1;'></center></div>";
        aretdiv[aretdivpos++]="<div style='background:wheat;position:absolute; top:45; left:0; overflow:scroll; overflow-x:hidden; width:200; height:108;border-bottom:2px solid black;'>";
        for(k=0; k<nsugestions; k++){
           sugestion= sugestions.item(k);
           aretdiv[aretdivpos++]="<div onmouseover='this.style.background=\"#ffffff\"' onmouseout='this.style.background=\"wheat\"' STYLE='font-family:verdana;FONT-SIZE: 8pt; height:20px; background:wheat; border:1px solid black; padding:3px; padding-left:10px; cursor:hand;' onclick='parent.processPopup(\""+sugestion.text.replace(/\'/gi,"&#39;")+"\");' id='div_sug_"+wordoffset+"' name='div_sug_"+wordoffset+"'>"+(k+1)+")&nbsp;"+sugestion.text.replace(/\'/gi,'&#39;')+"</div>";
        }
        if (nsugestions==0){
           aretdiv[aretdivpos++]="<div onmouseover='this.style.background=\"#ffffff\"' onmouseout='this.style.background=\"wheat\"' STYLE='font-family:verdana;FONT-SIZE: 8pt; height:20px; background:wheat; border:1px solid black; padding:3px; padding-left:10px; cursor:hand;' id='div_sug_"+wordoffset+"' name='div_sug_"+wordoffset+"' onclick='parent.drPopup.hide()'>Brak propozycji</div>";
        } else{
           aretdiv[aretdivpos++]="<div onmouseover='this.style.background=\"#ffffff\"' onmouseout='this.style.background=\"wheat\"' STYLE='font-family:verdana;FONT-SIZE: 8pt; height:20px; color:red; background:wheat; border:1px solid black; padding:3px; padding-left:10px; cursor:hand;' onclick='parent.processPopup(\""+original.replace(/\'/gi,"&#39;")+"\");' id='div_sug_"+wordoffset+"' name='div_sug_"+wordoffset+"'>"+(k+1)+') '+original.replace(/\'/gi,"&#39;")+"</div>";
        }   
        aretdiv[aretdivpos++]="</div>";
        aretdiv[aretdivpos++]="</div>";       
        apos=parseInt(offset)+linestart+original.length+aimgelementscount;
        s="<span id='myimgspan_"+aimgelementscount+"' name='myimgspan_"+aimgelementscount+"'><img id='myimg_"+aimgelementscount+"' name='myimg_"+aimgelementscount+"' nr='"+aimgelementscount+"' add='0' mOriginal='"+original+"' mLastOriginal='"+original+"' mOffset='"+wordoffset+"' mLen='"+original.length+"' mPos='"+apos+"' src='../images/SpellLetterRed.png' onclick='javascript:parent.processMTCommand(this);' style='cursor:hand;border-color:#ff0000' border=0></span>";
        if (atextrange.findText(original,1,6)){
           atextrange.pasteHTML(original+s);
           aimgelementscount++;
        }
        else {
          if (atextrange.findText(original,1,4)){
             atextrange.pasteHTML(original+s);
             aimgelementscount++;
          }
        }

     }
     linestart=linestart+parseInt(linelength)+1;
  }

  xmlHttp=null;
  xmlDoc=null;
  document.all('divmessage').style.visibility='hidden';  

  if (nmistakes==0){
     alert("Tekst poprawny");
     window.returnValue='';
     window.close();     
  }
  else{
     lastelement=aframedocument.all('myimg_0');
     document.all('idliczbabledow').innerHTML=nmistakes;
     findMistake(0);     
     document.all("divcontrols").disabled=false;
  }         

  document.all("correctioncache").innerHTML=aretdiv.join("");  
  }  
}
function doCorrectOK(){
  doClearImages();
  aframedocument=document.all('myspellframe').contentWindow.document;
  window.returnValue=aframedocument.body.innerHTML;
  window.close();
}
function doAdd(aword){
  var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
  xmlHttp.open('POST', 'SpellCheckXML.asp', false);
  sendtext='<DANE mode="add" language="'+alanguage+'"><![CDATA[' + aword.replace(/\&\#39\;/gi,"'").toLowerCase() + ']]></DANE>';
  xmlHttp.send(sendtext);
  var resp=xmlHttp.responseText;
  if (resp=='ERROR'){
     alert('Problem z serwerem!');
  } else{
     xmlDoc.async=false;
     xmlDoc.loadXML(xmlHttp.responseXML.xml);
     alert('Wyraz dodano do s³ownika');
     lastelement.setAttribute('add','1');
     lastelement.setAttribute('src','../images/SpellLetterGreen.png');     
  }
  xmlHttp=null;
  xmlDoc=null;
}
</script>

</head>
<body  onload='javascript:replacehtml();'>

<div style='position:absolute; top:0;height=100%'>
   <table border=0 height=100% width=100% cellpadding=5 cellspacing=0>
     <tr>
        <td align=left>
           <b>Liczba b³êdów: <span id='idliczbabledow'>0</span>, iloœæ s³ów: <span id='idliczbaslow'>0</span>, czas sprawdzania: <span id='idczas'>0</span> sek.</b>
        </td>   
        <td align=right><span id='idlang'></span></td>        
     </tr>
     <tr>
        <td height=100% colspan=2>
           <iframe name=myspellframe id=myspellframe class="normal" style="height:100%; width:100%;background-color:WINDOW; border:inset; overflow=auto;WHITE-SPACE:wrap;"></iframe>
        </td>
     </tr>
     <tr>
        <td align=center colspan=2>
           <div id='divcontrols' disabled='true'>
              <table border=0 height=100% width=100% cellpadding=5 cellspacing=0>     
                <tr>
                   <td align=center colspan=2>
                      <button onclick='javascript:findMistake(0)'>PIERWSZY</button>&nbsp;&nbsp;
                      <button onclick='javascript:findMistake(1)'>POPRZEDNI</button>&nbsp;&nbsp;
                      <button onclick='javascript:findMistake(2)'>NASTÊPNY</button>&nbsp;&nbsp;
                      <button onclick='javascript:findMistake(3)'>OSTATNI</button>
                   </td>
                </tr>
                <tr>
                   <td align=center colspan=2>Twoja propozycja dla b³êdu nr <span id='idnrbledu'>0</span>:&nbsp;
                      <input type=text name='lastclickedword'>&nbsp;&nbsp;
                      <button onclick='javascript:processPopup(document.all("lastclickedword").value)'>ZMIEÑ</button>
                   </td>
                </tr>
                <tr>
                   <td align=right width=50%>
                      <button onclick='javascript:doCorrectOK(document.all("myframe"))'><b>AKCEPTUJ</b></button>
                   </td>
                   <td align=left width=50%>
                      <button onclick="javascript:window.returnValue='';window.close();">ANULUJ</button>
                   </td>
                </tr>
                <tr>
                   <td colspan=2>&nbsp</td>
                </tr>
              </table>
           </div>
        </td>
     </tr>   
   </table>        
</div>
<div id='divmessage' style='position:absolute;width:300;'>
<TABLE cellpadding=3 cellspacing=0 bgcolor="yellow" style="border:solid 1px #235586"><TR><TD class=a0 align=center>
<font color="navy" size='+2'>Trwa sprawdzanie pisowni...</font>
</td></tr>
</Table>
</div>
<div id=correctioncache name=correctioncache></div>
</body>
</html>