<%@ CodePage=65001 LANGUAGE="VBSCRIPT" %><!--#include file ="definc_createIISI.asp"--><!DOCTYPE html>
<html lang="pl">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="format-detection" content="telephone=no">
    <meta name="msapplication-tap-highlight" content="no">
    <meta name="viewport" content="user-scalable=no,initial-scale=1,maximum-scale=1,minimum-scale=1,width=device-width">
    <title>ICORManager - start page</title>
    <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700|Material+Icons" rel="stylesheet" type="text/css">
    <link rel="icon" href="favicon.png" type="image/x-icon">
    <link href="https://cdn.jsdelivr.net/npm/quasar-framework@latest/dist/umd/quasar.mat.min.css" rel="stylesheet" type="text/css">
  </head>
  <body>
    <div id="q-app">
      <q-layout view="lHh Lpr fff">
         <!--
        <q-layout-header>
         -->
          <q-toolbar inverted>
            <img src="statics/icons/i_favicon-32x32bw.png">
            <q-toolbar-title>
              ICOR wita dziś!
              <div slot="subtitle">i jutro też będzie witał i pojutrze również</div>
            </q-toolbar-title>
          </q-toolbar>
         <!--
        </q-layout-header>
         -->
        <q-page-container>
          <my-page></my-page>
        </q-page-container>
      </q-layout>
    </div>

    <script type="text/x-template" id="my-page">
      <q-page padding>
       <q-list no-border link inset-delimiter multiline>
         <q-list-header>Najczęściej zadawane pytania:</q-list-header>
         <q-item @click.native="launch('http://www.microsoft.com/downloads/info.aspx?na=90&p=&SrcDisplayLang=pl&SrcCategoryId=&SrcFamilyId=7287252c-402e-4f72-97a5-e0fd290d4b76&u=http%3a%2f%2fdownload.microsoft.com%2fdownload%2fd%2f9%2fd%2fd9d4ac0f-f9bd-47f0-9c95-a81ff494b905%2fOWC11.EXE')">
           <q-item-side icon="chat"></q-item-side>
           <q-item-main label="Mam kłopoty z działaniem tabeli przestawnej, arkusza kalkulacyjnego" sublabel="Proszę zainstalować w trybie administracyjnym najnowszą wersję Office Windows Controls."></q-item-main>
         </q-item>
         <!--
         <q-item @click.native="launch('https://discord.gg/5TDhbDg')">
           <q-item-side icon="school"></q-item-side>
           <q-item-main label="Discord Chat Channel" sublabel="https://discord.gg/5TDhbDg"></q-item-main>
         </q-item>
         <q-item @click.native="launch('http://forum.quasar-framework.org')">
           <q-item-side icon="forum"></q-item-side>
           <q-item-main label="Forum" sublabel="forum.quasar-framework.org"></q-item-main>
         </q-item>
         <q-item @click.native="launch('https://twitter.com/quasarframework')">
           <q-item-side icon="rss feed"></q-item-side>
           <q-item-main label="Twitter" sublabel="@quasarframework"></q-item-main>
         </q-item>
         -->
       </q-list>
      </q-page>
    </script>
    
    <script src="https://cdn.jsdelivr.net/npm/vue@latest/dist/vue.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/quasar-framework@latest/dist/umd/quasar.mat.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/quasar-framework@latest/dist/umd/i18n.pl.umd.min.js"></script>
    <script src="https://unpkg.com/axios/dist/axios.min.js"></script>

    <script>
      Quasar.i18n.set(Quasar.i18n.pl)
      
      Vue.component('my-page', {
        template: '#my-page',
        methods: {
          launch: function (url) {
            Quasar.utils.openURL(url)
          }
        }
      })

      new Vue({
        el: '#q-app',
        data: function () {
          return {
          }
        },
        mounted: function () {
        },
        methods: {
          launch: function (url) {
            Quasar.utils.openURL(url)
          }
        }
      })
    </script>
  </body>
</html>
