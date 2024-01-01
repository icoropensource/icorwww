<%@ CodePage=65001 LANGUAGE="VBSCRIPT" %><!--#include file ="definc_createIISI.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<!--
<link rel=STYLESHEET type="text/css" href="/icormanager/icor.css" title="SOI">
-->
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta name="description" content="Opis strony">
<meta name="pragma" content="no-cache">
<meta name="keywords" content="ICOR object oriented database WWW information management repository">
<meta name="author" content="<%=Application("META_AUTHOR")%>">
<script type="text/javascript" src="/icorlib/jquery/jquery-latest.min.js"></script>
<link rel="stylesheet" type="text/css" href="/icorlib/jquery/plugins/ui-1.8.14/css/<%=Dession("UI_SKIN")%>/jquery-ui-1.8.14.custom.min.css">
<script type="text/javascript" src="/icorlib/jquery/plugins/ui-1.8.14/js/jquery-ui-1.8.14.custom.js"></script>
<style type="text/css">
  img.logo_mp {border:none;float:left}
  img.logo_icor {border:none;float:right}
  img.powered_by {border:none;margin-top:5px;}
  div.stopka {text-align:center;border-top:solid 1px silver;margin-top:50px;}

html {}
body {padding:0px;margin:0px;
      font-size:11px; font-family:'Arial CE', Arial, sans-serif;}
/*
  h1.tytul {margin:0px;padding:3px;
            font-size:12px;font-weight:normal;font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;
            background-image:url('img/tlo_tytul.gif');background-position:0% 50%;
            border-bottom:solid 1px #2694e8}
  h1 {margin:0px;padding:0px;
      font-size:24px;font-weight:normal;color:#0061c5;
      font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;}
  h2 {font-size:20px;font-weight:normal;
      font-family:'Trebuchet MS','Arial CE', Arial, sans-serif;}

a {color:blue}
a:hover {color:red} 
a img {border:none}
a.reflistoutnavy {margin:3px;padding:5px;display:inline-block;width:auto;
                  font-size:10px;font-weight:bold;color:black;text-decoration:none;
                  background-image:url('img/tlo_tytul2.gif');background-position:0% 50%;
                  border:solid 1px #cccccc;}
a.reflistoutnavy:hover {color:black;background-image:url('img/tlo_tytul.gif');border:solid 1px #2694e8;}

div.tresc {width:99%;height:200px;padding:3px;
           overflow:scroll;overflow-x:hidden;
           background-color:#f6f6f6;border:solid 1px #cccccc}
div.opisy {padding:10px}

table.objectinfo {margin-bottom:20px;margin-top:20px;}
td.objectseditsysteminfo {width:250px;font-weight:bold;line-height:110%}
td.objectseditsysteminfovalue {width:250px;}

table.log_table {font-size:11px; font-family:'Arial CE', Arial, sans-serif;}
  table.log_table tr {}
  table.log_table td.objectseditdatafieldnameeven {padding:3px;border-bottom:solid 1px #aaaaaa;
                                                   background-color:#f6f6f6}
  table.log_table td.objectseditdatafieldnameodd {padding:3px;background-color:#ffffff;
                                                  border-bottom:solid 1px #aaaaaa;}

span.objectsviewcaption {font-weight:bold}
table.sort-table {margin-top:10px;
                  border:solid 1px #aaaaaa;
                  font-size:11px;font-family:'Arial CE', Arial, sans-serif;}
  table.sort-table thead {background-image:url('img/tlo_tytul2.gif');background-position:0% 50%;}
  table.sort-table thead tr {}
  table.sort-table thead td {padding:3px;border-right:solid 1px #aaaaaa;border-bottom:solid 1px #aaaaaa}
  table.sort-table tbody {}
  table.sort-table tbody tr {}
  table.sort-table tbody tr:hover {color:red}
  table.sort-table tbody td {padding:3px;}
  table.sort-table tbody td a {text-decoration:none}


.bold {font-weight:bold}
.left {text-align:left}
.right {text-align:right}
.flow_right {float:right}
.flow_left {float:left}
.clear {clear:both}
*/
  </style>
<title>Main page</title>
</head>
<BODY>
<!--
   <a target=new href="http://www.icor.pl"><IMG SRC="images/logo_icor.gif" class="logo_icor"></a>
   <a target=new href="http://www.mikroplan.com.pl"><IMG SRC="images/logo_mp.gif" class="logo_mp"></a>
 -->
<div id="mdiv1" class="ui-widget">
<div class="ui-widget-header">
    <h1>ICOR wita dziś!</h1>
</div>

    <h2>Najczęściej zadawane pytania</h2>
<OL>
  <LI><b>Czasami nie mogę zalogować się do ICORManager!</b><BR>
      Proszę następnym razem zwrócić uwagę na wciśnięty klawisz <i>Caps Lock</i>. Małe i duże litery w hasłach <U>są rozróżniane</U>.<BR><BR></li>
  <li><b>Mam kłopoty z działaniem tabeli przestawnej, arkusza kalkulacyjnego.</b><BR>
      Proszę zainstalować w trybie administracyjnym najnowszą wersję OWC. Odpowiednio, gdy na komputerze jest zainstalowy pakiet <a target="_new" href="http://www.microsoft.com/downloads/info.aspx?na=90&p=&SrcDisplayLang=pl&SrcCategoryId=&SrcFamilyId=982b0359-0a86-4fb2-a7ee-5f3a499515dd&u=http%3a%2f%2fdownload.microsoft.com%2fdownload%2fb%2ff%2f3%2fbf3f6bde-be0b-4afe-85f2-02e374030881%2fowc10.exe">MS Office XP</a> lub <a target="_new" href="http://www.microsoft.com/downloads/info.aspx?na=90&p=&SrcDisplayLang=pl&SrcCategoryId=&SrcFamilyId=7287252c-402e-4f72-97a5-e0fd290d4b76&u=http%3a%2f%2fdownload.microsoft.com%2fdownload%2fd%2f9%2fd%2fd9d4ac0f-f9bd-47f0-9c95-a81ff494b905%2fOWC11.EXE">MS Office 2003 lub nowszy</a>.<BR><BR></li>
</OL>
</div>


</BODY></HTML>
