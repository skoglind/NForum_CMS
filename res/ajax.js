var xmlhttp=false;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp = false;
   }
  }
@end @*/
if(!xmlhttp && typeof XMLHttpRequest != 'undefined'){
  xmlhttp = new XMLHttpRequest();
}

var xmlhttp_db=false;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp_db = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp_db = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp_db = false;
   }
  }
@end @*/
if(!xmlhttp_db && typeof XMLHttpRequest != 'undefined'){
  xmlhttp_db = new XMLHttpRequest();
}

var xmlhttp_ka=false;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp_ka = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp_ka = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp_ka = false;
   }
  }
@end @*/
if(!xmlhttp_ka && typeof XMLHttpRequest != 'undefined'){
  xmlhttp_ka = new XMLHttpRequest();
}

function getPage(sPage) {
  var xmlRet;
  
  xmlhttp.open("GET", sPage);
    xmlhttp.onreadystatechange = function(){
      if(xmlhttp.readyState == 4 && xmlhttp.status == 200){
        xmlRet = xmlhttp.responseText;
     
        if(xmlRet != "") {
          document.getElementById("statusarea").innerHTML = xmlRet;
        }
      }
    }
    xmlhttp.send(null);
}

function getPage_ka(sPage) {
  var xmlRet;
  
  xmlhttp_ka.open("GET", sPage);
    xmlhttp_ka.onreadystatechange = function(){
      if(xmlhttp_ka.readyState == 4 && xmlhttp_ka.status == 200){
        xmlRet = xmlhttp_ka.responseText;
      }
    }
    xmlhttp_ka.send(null);
}

function getPage_db(sPage,o) {
  var xmlRet;
  
  xmlhttp_db.open("GET", sPage);
    xmlhttp_db.onreadystatechange = function(){
      if(xmlhttp_db.readyState == 4 && xmlhttp_db.status == 200){
        xmlRet = xmlhttp_db.responseText;
        
        if(xmlRet != "") {
          document.getElementById(o).innerHTML = xmlRet;
        }
      }
    }
    xmlhttp_db.send(null);
}


function StayOnline() {
  getPage_ka("/_keepalive.asp");
  setTimeout("StayOnline();", 60000);
}