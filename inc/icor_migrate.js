// NAVBAR
function getParentFrame_old(aname) {
  var w = window, f, i = 7;
  while (i > 0) {
    f = w.frames[aname];
    if (f && (f.name == aname)) {
      return f;
    }
    w = w.parent;
    i = i - 1;
  }
  return undefined;
}

var tabdict = {};
var lasttabcid = -1;
var lasttabnr = -1;
var lasttabid = -1;

if (tabdict != null) {
  tabdict["ICOR"] = "ICOR";
}

var maxHappySound = 6;
var maxSadSound = 5;

var tabstate = {};

dParentFrame = {
  tablesheets : {
  },
  registerStateBadOK: function (astate) {

  },
  SetSheetForCid: function (cid, atoframe, atabid) {
    if (tabdict != null) {
      var lasttab = new Object();
      lasttab.cid = cid;
      lasttab.nr = atoframe;
      lasttab.id = atabid;
      tabdict[cid] = lasttab;
    } else {
      lasttabcid = cid;
      lasttabnr = atoframe;
      lasttabid = atabid;
    }
  },
  GetSheetForCid: function (cid) {
    if (tabdict != null) {
      if (cid in tabdict) {
        lasttab = tabdict[cid];
        return lasttab;
      }
    } else {
      if (lasttabcid == cid) {
        var newtab = new Object();
        newtab.cid = lasttabcid;
        newtab.nr = lasttabnr;
        newtab.id = lasttabid;
        return newtab;
      }
    }
  },
  displayMessage: function (s) {
    alert(s);
  },
  StartProgress: function (aname) {
    var ProgressURL = 'progresswindow.asp?name=' + aname
    var v = window.open(ProgressURL, '_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')
  },
  playHappySound: function () {
    var anum = Math.random();
    anum = anum * maxHappySound;
    anum = 1 + Math.floor(anum);
    try {
      var x = document.all("happySound" + anum);
      x.play();
    } catch (e) {
      document.all("sounds").insertAdjacentHTML("beforeEnd", "<embed src='inc/snd/success/happy" + anum + ".wav' hidden=true autostart=true loop=false id='happySound" + anum + "' name='happySound" + anum + "' VOLUME=100>");
    }
  },
  playSadSound: function () {
    var anum = Math.random();
    anum = anum * maxSadSound;
    anum = 1 + Math.floor(anum);
    try {
      var x = document.all("sadSound" + anum);
      x.play();
    } catch (e) {
      document.all("sounds").insertAdjacentHTML("beforeEnd", "<embed src='inc/snd/sad/sad" + anum + ".wav' hidden=true autostart=true loop=false id='sadSound" + anum + "' name='sadSound" + anum + "' VOLUME=100>");
    }
  },
  stateChecked: function (oXmlDoc, astateid, atext) {
    if (oXmlDoc == null || oXmlDoc.documentElement == null) {
      if (oXmlDoc == null) {
        alert("State oXmlDoc (1) == null");
      }
    } else {
      var root = oXmlDoc.documentElement;
      var cs = root.childNodes;
      var l = cs.length;
      var aname = "";
      var avalue = "";
      for (var i = 0; i < l; i++) {
        if (cs[i].tagName == "STATE") {
          aname = cs[i].getAttribute("name");
          avalue = cs[i].getAttribute("value");
        }
      }
      if (avalue == "DEL") {
        try {
          delete tabstate[astateid];
        } catch (ex) {; }
      }
      if (avalue == "OK") {
        delete tabstate[astateid];
        window.focus();
        dParentFrame.playHappySound();
        //        document.all("stateinfo").stop();
        //        document.all("stateinfo").style.display='none';
        alert(aname);
      }
      if (avalue == "BAD") {
        delete tabstate[astateid];
        window.focus();
        dParentFrame.playSadSound();
        //        document.all("stateinfo").stop();
        //        document.all("stateinfo").style.display='none';
        alert(aname);
      }
      if (avalue.slice(0, 1) == '#') {
        //        document.all("stateinfo").innerHTML=avalue.slice(1);
        //        document.all("stateinfo").style.display='';
        //        document.all("stateinfo").start();
      }
      //      if (avalue=="RUN") {
      //        tabstate.Remove(astateid);
      //      }
    }
  },
  stateChecker: function stateChecker() { // $$ do zmiany dla FF
    var astateid, w;
    w = 0;
    for (astateid in tabstate) {
      w = 1;
      var xmlHttp = XmlHttp.create();
      xmlHttp.open("GET", "icorsync.asp?mode=get&state=" + astateid, true);
      xmlHttp.onreadystatechange = function () {
        if (xmlHttp.readyState == 4) {
          dParentFrame.stateChecked(xmlHttp.responseXML, astateid, xmlHttp.responseText);
        }
      };
      try {
        window.setTimeout(function () {
          try {
            xmlHttp.send(null);
          } catch (ex) {;
          }
        }, 10);
      } catch (ex) {;
      }
    }
    if (w == 1) {
      window.setTimeout(dParentFrame.stateChecker, 7000);
    }
  },
  registerStateBadOK: function (aid) {
    // tabstate[aid] = aid;
    // window.setTimeout(dParentFrame.stateChecker, 7000);
    window.parent.postMessage({
      type: 'registerStateBadOK',
      aid: aid
    }, '*');
  },
  '_': '_'
};

function getParentFrame(aname) {
  return dParentFrame;
}

// UTIL
function htmlDecode(str) {
  var entMap = { 'quot': 34, 'amp': 38, 'apos': 39, 'lt': 60, 'gt': 62 };
  return str.replace(/&([^;]+);/g, function (m, n) {
    var code;
    if (n.substr(0, 1) == '#') {
      if (n.substr(1, 1) == 'x') {
        code = parseInt(n.substr(2), 16);
      } else {
        code = parseInt(n.substr(1), 10);
      }
    } else {
      code = entMap[n];
    }
    return (code === undefined || code === NaN) ? '&' + n + ';' : String.fromCharCode(code);
  });
}
