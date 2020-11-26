function hide(o) {
  document.getElementById(o).style.display = "none";
}

function show(o) {
  document.getElementById(o).style.display = "block";
}

function removeMe(m,c) {
  var holderDiv = document.getElementById(m);
  var oldDiv = document.getElementById(c);
  holderDiv.removeChild(oldDiv);
  
  if(getSelectedRadio("vStandardTitel")==-1) {
    setSelectedRadio("vStandardTitel");
  }
}

function getSlumpID() {
  var randomnumber = Math.floor(Math.random()*499999) + 2500000;
  return randomnumber;
}

function addMe(m,id,titel,extra,release,regionskod,utgivare,utgivaretext,region) {
  var holderDIV = document.getElementById(m);
  var hiddenDIV = document.getElementById("t_hiddenbox");
  var newDIV = document.createElement("div");
  
  newHTML = hiddenDIV.innerHTML;
  newHTML = newHTML.replace(/mmID/g, id);
  newDIV.innerHTML = newHTML;
  
  newDIV.setAttribute("id","TitelRow_" + id);
  newDIV.className = "in_row titelbox";
  holderDIV.appendChild(newDIV);
  
  document.getElementById("vTitel_" + id).value = titel;
  document.getElementById("vExtra_" + id).value = extra;
  document.getElementById("vRelease_" + id).value = release;
  document.getElementById("vRegionskod_" + id).value = regionskod;
  document.getElementById("vUtgivareID_" + id).value = utgivare;
  document.getElementById("vUtgivareText_" + id).value = utgivaretext;
  
  if(getSelectedRadio("vStandardTitel")==-1) {
    setSelectedRadio("vStandardTitel");
  }
}

function selectValue(o,v) {
  var er = document.getElementById(o);
  for(var i=0; i < er.options.length; i++) {
    if(er.options[i].value == v) {
      er.options[i].selected = true;
    }
  }
}

function hidenames(o) {
  var e = document.getElementsByName(o);
  for(var i=0;i<e.length;i++){
    e[i].style.display = "none";
  }
}

function shownames(o) {
  var e = document.getElementsByName(o);
  for(var i=0;i<e.length;i++){
    e[i].style.display = "block";
  }
}

function hideclass(c) {
  var tabTmp = new Array();
  tabTmp = document.getElementsByTagName("*");
  for (i=0; i<tabTmp.length; i++) {
    if (tabTmp[i].className==c) {
      tabTmp[i].style.display = "none";
    }
  }
}

function showclass(c) {
  var tabTmp = new Array();
  tabTmp = document.getElementsByTagName("*");
  for (i=0; i<tabTmp.length; i++) {
    if (tabTmp[i].className==c) {
      tabTmp[i].style.display = "block";
    }
  }
}

function toggle(o) {
  var e = document.getElementById(o);
  if(e.checked) {
    e.checked = false;
  } else {
    e.checked = true;
  }
}

function toggletrue(o) {
  var e = document.getElementById(o);
  e.checked = true;
}

function getToggle(snamn) {
  var n = document.getElementsByName(snamn);
  var r = document.getElementById(snamn);
  var anyChecked; 

  if(r.checked) {
    for(var i=0;i<n.length;i++){
      n[i].checked = false;
    }
    
    r.checked = true;
  }
}

function setToggle(snamn) {
  var n = document.getElementsByName(snamn);
  var r = document.getElementById(snamn);
  var anyChecked; 

  for(var i=0;i<n.length;i++){
    if(n[i].checked) {
      anyChecked = true;
    }
  }
  
  if(!anyChecked) {
    r.checked = true;
  } else {
    r.checked = false;
  }
}

function infobox(b) {
  if(b) {
    show("informationbox");
  } else {
    hide("informationbox");
  }
}

function infobox_text(s, b2, b2text, b2action) {
  nStr = '<h2>Meddelande</h2>' + s + '<div class="buttons"><input type="button" value="Stäng" onclick="infobox(false);">';
  if(b2) {
    nStr = nStr + '<input type="button" value="' + b2text + '" style="font-weight: bold;" onclick="' + b2action + '"></div>';
  } else {
    nStr = nStr + '</div>'
  }

  document.getElementById("infobox_text").innerHTML = nStr;
}

function mkdisable(o,b) {
  document.getElementById(o).disabled = b;
}

function cpVal(o) {
  if(document.getElementById(o)) {
    if(document.getElementById(o).value=="YES") {
      if(document.getElementById(o).checked) {
        document.getElementsByName(o + "_cp")[0].value = "YES";
      } else {
        document.getElementsByName(o + "_cp")[0].value = "";
      }
    } else {
      document.getElementsByName(o + "_cp")[0].value = document.getElementById(o).value;
    }
  }
}

function addText(o,s) {
  var obj = document.getElementById(o);
  obj.focus();
  
  if (document.selection) {
    var sel = document.selection.createRange();
    if(sel.text.length > 0) {
      sel.text = "[" + s + "]" + sel.text + "[/" + s + "]";
    } else {
      stxt = "[" + s + "][/" + s + "]";
      obj.value = obj.value + stxt;
    }
  } else {
    lStart = obj.selectionStart;
    lEnd = obj.selectionEnd;
    if(lEnd > lStart) {
      tFront = obj.value.substr(0, lStart);
      sSel = obj.value.substr(lStart, (lEnd - lStart));
      tBack = obj.value.substr(lEnd, (obj.value.length - lEnd));
      obj.value = tFront + "[" + s + "]" + sSel + "[/" + s + "]" + tBack;
    } else {
      stxt = "[" + s + "][/" + s + "]";
      obj.value = obj.value + stxt;
    }
  }
}

function addTextEnd(o,s) {
  var obj = document.getElementById(o);
  obj.focus();

  obj.value = obj.value + s;
}

function getSelectedRadio(o) {
  var n = document.getElementsByName(o);
  var k;

  for(var i=0;i<n.length;i++){
    if(n[i].checked) {
      k = n[i].value;
    }
  }
  
  if(!k) {k=-1}
  
  return k;
}

function setSelectedRadio(o) {
  var n = document.getElementsByName(o);
  if(n.length<2) {
    addMe('titles',getSlumpID(),'','','','',0,'Ingen utgivare vald','');
  }
  
  n[0].checked = true;
}

function setRadio(o,v) {
  var n = document.getElementsByName(o);
  for(var i=0; i<n.length; i++) {
    if(n[i].value == v) {
      n[i].checked = true;
    }
  }
}

function clearPicker(fld1, fld2, txt) {
  document.getElementById(fld1).value = "0";
  document.getElementById(fld2).value = txt;
}

var eBoxID;
var eBoxText

function showPicker(sPage,eID,eText){
  eBoxID   = eID;
  eBoxText = eText;
  nWin = window.open(sPage, 'PICKER','toolbar=no, width=300, height=375, location=no, menubar=no');
  if(!nWin) {
    alert('Du måste tillåta popup-fönster på denna sida för att kunna använda systemet!');
  } else {
    nWin.focus();
  }
}

function setValue(vID, vText) {
  document.getElementById(eBoxID).value = vID;
  document.getElementById(eBoxText).value = vText;
}

function sendbackdata(vID) {
  var e = document.getElementById("ibox" + vID);
  vText = e.value;
  opener.setValue(vID,vText);
  window.close();
}

function mkdisable_obj(o,b) {
  var e = document.getElementsByName(o);
  for(var i=0;i<e.length;i++){
    e[i].disabled = b;
  }
}

function fldReset(o,s) {
  document.getElementById(o).value = s;
}

function doSubmit(e,q) {
  if(q != "a=") {
    document.getElementById(e).action = "__do.asp?"+ q;
    document.getElementById(e).submit();
  } else {
    alert("Inget alternativ valt!");
  }
}

function ajax_loader(b,s) {
  if(b) {
    document.getElementById("do_what").innerHTML = s;
    
    show("ajax_loading");
    hide("ajax_waiting");
  } else {
    document.getElementById("dont_do").innerHTML = s;
    
    show("ajax_waiting");
    hide("ajax_loading");
  }
}

function btnHandle(b) {
  mkdisable_obj("savebtn",b);
  mkdisable_obj("saveimg",b);
  mkdisable_obj("undoimg",b);
  mkdisable_obj("uplbtn",b);
}

function saveform(e,i) {
  btnHandle(true);
  
  switch(i)
  {
    case 0:
      q = "a=save";
    break;  
    case 1:
      q = "a=save&ea=continue";
    break;
    case 2:
      q = "a=save&ea=return";
    break;
  }
  
  document.getElementById(e).action = "__do.asp?"+ q;
  document.getElementById(e).target = "processbox";
  document.getElementById(e).submit();
  
  ajax_loader(true, "Sparar data...");
}

function doform(e,ob) {
  btnHandle(true);
  
  q = "a=" + ob;
  
  document.getElementById(e).action = "__do.asp?"+ q;
  document.getElementById(e).target = "processbox";
  document.getElementById(e).submit();
  
  ajax_loader(true, "Utför åtgärd...");
}

function savefailed(failmess) { 
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
  
  infobox_text(failmess, false, '', '');
  infobox(true);
}

function savefinished(sd,id,b,l) { 
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
  
  document.getElementById("vID").value = id;
  
  if(b) {
    setvar = "location.href='" + l + "';";
    
    infobox_text('<p>Allt är nu sparat och du kan välja att utföra nästa åtgärd.</p>', true, 'Fortsätt...', setvar);
    infobox(true);
  }
  
  local_ResetFields();
  getPage("__innerfld.asp?e=" + id);
  
  document.getElementById("savedstatus").style.backgroundImage = "url(/design/icons/radio_true.png)";
  document.getElementById("savedstatus").innerHTML = sd;
}

function dofinished(mess) { 
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
  
  window.alert(mess);
}

function deleteimg(id,objid,area) {
  frames["processbox"].location.href = "/_imgupl.asp?m=2&vid=" + id + "&vobjid=" + objid + "&varea=" + area;
  
  deletewaiting();
}

function deletewaiting() {
  btnHandle(true);
  ajax_loader(true, "Raderar bild...");
}

function deletefailed() {
  infobox_text('<p>Bilden kunde inte raderas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Ingen bild</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function deletefinished(id) {
  //show('f_textmess');
  hide('f_new');
  mkdisable('btnew', false);
  btnHandle(false);
  
  hide("f_id" + id);
  
  ajax_loader(false, "Avvaktar.");
}

function uploadimg(id) {
  var fldName;
  if(id==0) {
    fldName = "imgupl_new";
  } else {
    fldName = "imgupl_id" + id;
  }
    
  document.getElementById(fldName).action = "/_imgupl.asp?m=1";
  document.getElementById(fldName).target = "processbox";
  document.getElementById(fldName).submit();
  
  uploadwaiting();
}

function uploadwaiting() {
  btnHandle(true);
  ajax_loader(true, "Sparar bild...");
}

function uploadfailed() {
  infobox_text('<p>Bilden kunde inte sparas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Större än 2MB<br>- Ingen bild<br>- Fel format (inte png, jpg/jpeg, bmp eller gif).</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function uploadfinished(isnew, id, imgname, imgtext) {
  //show('f_textmess');
  hide('f_new');
  mkdisable('btnew', false);
  btnHandle(false);
  
  addHTML = document.getElementById("f_hiddenbox").innerHTML;
  addHTML = addHTML.replace(/\'%ID%\'/g, id);
  addHTML = addHTML.replace(/%ID%/g, id);
  addHTML = addHTML.replace(/%THAIMG%/g, imgname);
  
  var sBoxes = imgtext.split("||");
  for(i = 0; i < sBoxes.length; i++){
    if(sBoxes[i]!="") {
      addHTML = addHTML.replace(new RegExp("%" + sBoxes[i] + "%", "gi"), sBoxes[i+1]);
      i++;
    }
  }
  
  if(isnew) {
    allHTML = document.getElementById("imgholder").innerHTML;
    allHTML = allHTML + "<div class='imgblock' id='f_id" + id + "'>" + addHTML + "</div>";
    document.getElementById("imgholder").innerHTML = allHTML;
  } else {
    document.getElementById("f_id" + id).innerHTML = addHTML;
  }
  
  ajax_loader(false, "Avvaktar.");
}

function smalldeleteimg(id,area) {
  frames["processbox"].location.href = "/_imgupl.asp?m=4&lid=" + id + "&sarea=" + area;
  
  btnHandle(true);
  smalldeletewaiting();
}

function smalldeletewaiting() {
  ajax_loader(true, "Raderar bild...");
}

function smalldeletefailed() {
  infobox_text('<p>Bilden kunde inte raderas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Ingen bild</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function smalldeletefinished(id) {
  btnHandle(false);
  getPage("__innerfld.asp?e=" + id);
  ajax_loader(false, "Avvaktar.");
}

function smalluploadimg(fldName) {   
  document.getElementById(fldName).action = "/_imgupl.asp?m=3";
  document.getElementById(fldName).target = "processbox";
  document.getElementById(fldName).submit();
  btnHandle(true);
  smalluploadwaiting();
}

function smalluploadwaiting() {
  ajax_loader(true, "Sparar bild...");
}

function smalluploadfailed() {
  infobox_text('<p>Bilden kunde inte sparas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Större än 2MB<br>- Ingen bild<br>- Fel format (inte png, jpg/jpeg, bmp eller gif).</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function smalluploadfinished(id) {
  btnHandle(false);
  getPage("__innerfld.asp?e=" + id);
  ajax_loader(false, "Avvaktar.");
}

function gluesave(id,objid,area) {
  btnHandle(true);
  
  document.getElementById("glueform").action = "__do.asp?a=glue";
  document.getElementById("glueform").target = "processbox";
  document.getElementById("glueform").submit();
  
  glueuploadwaiting();
}

function gluewaiting() {
  btnHandle(true);
  ajax_loader(true, "Slår ihop...");
}

function gluefailed() {
  infobox_text('<p>Kunde inte slås ihop:<br>- Du saknar behörighet<br>- Objekt att slå ihop med är inte valt</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function gluefinished(l) {
  btnHandle(false);
  
  ajax_loader(false, "Avvaktar.");
  
  setvar = "location.href='" + l + "';";
    
  infobox_text('<p>Nu är de ihopslagna och aktuellt objekt är borttaget, välj fortsätt för att gå till det nya objektet.</p>', true, 'Fortsätt...', setvar);
  infobox(true);
}

function boxuploadimg() {
  btnHandle(true);
  
  document.getElementById("boxart_upl").action = "/_imgupl.asp?m=5";
  document.getElementById("boxart_upl").target = "processbox";
  document.getElementById("boxart_upl").submit();
  
  boxuploadwaiting();
}

function boxuploadwaiting() {
  ajax_loader(true, "Sparar bild...");
}

function boxuploadfailed() {
  infobox_text('<p>Bilden kunde inte sparas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Större än 2MB<br>- Ingen bild<br>- Fel format (inte png, jpg/jpeg, bmp eller gif).</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function boxuploadfinished(id,art,file) {
  var e = document.getElementById("upload_pic_" + art + "_" + id);
  e.src = file;
  
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function boxdeleteimg(id,art,area) {
  frames["processbox"].location.href = "/_imgupl.asp?m=6&lid=" + id + "&sart=" + art + "&sarea=" + area ;
  
  btnHandle(true);
  boxdeletewaiting();
}

function boxdeletewaiting() {
  ajax_loader(true, "Raderar bild...");
}

function boxdeletefailed() {
  infobox_text('<p>Bilden kunde inte raderas! Felet kan bero på följande orsaker:<br>- Du saknar behörighet<br>- Ingen bild</p>', false, '', '');
  infobox(true);
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function boxdeletefinished(id,art) {
  var e = document.getElementById("upload_pic_" + art + "_" + id);
  e.src = "/design/noimg.gif";
  btnHandle(false);
  ajax_loader(false, "Avvaktar.");
}

function setBoxart(id,bx_Fram,bx_Bak,bx_Man,bx_Kas) {
  document.getElementById("upload_pic_1_" + id).src = bx_Fram;
  document.getElementById("upload_pic_2_" + id).src = bx_Bak;
  document.getElementById("upload_pic_3_" + id).src = bx_Man;
  document.getElementById("upload_pic_4_" + id).src = bx_Kas;
}

function showPic(e) {
  winP = window.open(e.replace("w=80&h=80","w=800&h=600"),"BigPic","width=820, height=620");
  if(!winP) {
    alert("Du måste tillåta popup i din webbläsare för att förhandsgranska bilderna!");
  } else {
    winP.focus();
  }
}

function hide_uploadbox() {
  var o = document.getElementById("uploadbox");
  o.style.display = "none";
}

function show_uploadbox(btn,id,art,text,game) {
  var o = document.getElementById("uploadbox");
  var e = document.getElementById(btn);
  
  o.style.left = (getPageOffsetLeft(e) - 64) + "px";
  o.style.top  = (getPageOffsetTop(e) - 6) + "px";
  
  document.getElementById("upload_ID").value = id;
  document.getElementById("upload_Art").value = art;
  document.getElementById("game_ID").value = game;
  document.getElementById("upload_File").value = "";
  document.getElementById("upload_Text").innerHTML = text;
  
  o.style.display = "block";
}

var inWait;
var timern;
var ttText;
var ttKonsol;
var ttID;
var ttSpel;

function reduceLoad(text,konsol,id,spel) {
  if(timern) {clearTimeout(timern);}
  
  ttText    = text;
  ttKonsol  = konsol;
  ttID      = id;
  ttSpel    = spel;
  
  timern = setTimeout("checkTitleExists();", 500);
}

function checkTitleExists() {
  clearTimeout(timern);
  
  var e = document.getElementById(ttText);
  var k = document.getElementById("titlechecker");
  var t;
  
  chkText = e.value;
  
  if(chkText.length < 5) {
    k.innerHTML = "";
    k.style.display = "none";
  } else {
    k.style.left = (getPageOffsetLeft(e) + 2) + "px";
    k.style.top  = (getPageOffsetTop(e) + 22) + "px";
    
    getPage_db("/cms_FindTitle.asp?t=" + escape(chkText) + "&k=" + ttKonsol + "&e=" + ttID + "&s=" + ttSpel, "titlechecker");
    
    k.style.display = "block";
  }
}

function hideTitleExists() {
  var k = document.getElementById("titlechecker");
  
  k.innerHTML = "";
  k.style.display = "none";
}

function getPageOffsetLeft (el) {
  var ol=el.offsetLeft;
  while ((el=el.offsetParent) != null) { ol += el.offsetLeft; }
  return ol;
}

function getPageOffsetTop (el) {
  var ot=el.offsetTop;
  while ((el=el.offsetParent) != null) { ot += el.offsetTop; }
  return ot;
}

function addRow(id) {
  var holderDIV = document.getElementById("allrows");
  var hiddenDIV = document.getElementById("rowclone");
  var newDIV = document.createElement("div");
  
  newHTML = hiddenDIV.innerHTML;
  newHTML = newHTML.replace(/XXXX/g, id);
  newDIV.innerHTML = newHTML;
  
  newDIV.setAttribute("id","Row" + id);
  newDIV.className = "rad";
  holderDIV.appendChild(newDIV);
}

function delRow(id) {
  var holderDiv = document.getElementById("allrows");
  var oldDiv = document.getElementById("Row" + id);
  holderDiv.removeChild(oldDiv); 
}

function setRowData(id,f,data) {
  document.getElementById(f + "_" + id).value = data;
}
