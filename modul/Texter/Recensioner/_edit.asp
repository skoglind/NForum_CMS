<% 
  cON_PAGE = "Hantera recension - Recensioner - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS1") Then Response.Redirect("/")
  %>

  <%
  lID = Request.QueryString("e")
  If Not IsNumeric(lID) Or lID = Empty Then lID = 0
  lID = CLng(lID)
  
  ' #### BEHÖRIGHET ####
  If NOT GetAcc("CMS111") Then sBFilter = " AND rSkapadAv = " & cCMS_ID & " AND NOT rStatus = 0"
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT *, anvDB1.aNamn AS Anv1, anvDB2.aNamn AS Anv2 " & _
          "FROM (cms_Recensioner " & _ 
          "LEFT JOIN fsBB_Anv AS AnvDB1 ON cms_Recensioner.rSkapadAv = AnvDB1.aID) " & _
          "LEFT JOIN fsBB_Anv AS AnvDB2 ON cms_Recensioner.rPubliceradAv = AnvDB2.aID " & _
          "LEFT JOIN cms_Spel ON sID = rSpelID " & _
          "LEFT JOIN cms_SpelTitlar ON tID = sStandard_Titel " & _
          "WHERE rID = " & CLng(lID) & sBFilter
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Kategori           = 1
      cADD_Betyg              = 0
      
      cADD_SpelID             = 0
      cADD_SpelText           = "Inget spel valt"
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("rID")
      cADD_Kategori           = rsDB("rKategori")
      cADD_Titel              = sEncode(rsDB("rTitel"))
      cADD_Text               = sEncode(rsDB("rText"))
      cADD_Notes              = sEncode(rsDB("rNotes"))
      cADD_Short              = sEncode(rsDB("rShort"))
      cADD_Nyckelord          = sEncode(rsDB("rNyckelord"))
      cADD_Betyg              = CLng(rsDB("rBetyg"))
      
      cADD_bAnvRec            = rsDB("rAnvandarRec")
      
      cADD_SpelID             = CLng(rsDB("rSpelID"))
      If cADD_SpelID = 0 Then
        cADD_SpelText           = "Inget spel valt"
      Else
        cADD_SpelText           = lstKonsolXShort(rsDB("sKonsol")) & " | " & rsDB("tTitel")
      End If
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida
  ' ##################
  
  If Not rsDB.EOF Then If Not GetAcc("CMS11") And rsDB("rStatus") <> 1 Then bCantSave = True
  %>
  
  <script type="text/javascript">
    function cpFlds() {
      cpVal('vNotes');
      cpVal('vStatus');
      cpVal('vPublTyp');
      cpVal('vPublDatum');
      cpVal('vNySkapare');
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Hantera recension</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Konsol</div>
      <div class="input">
        <select name="vKategori">
          <% For zx = 1 To lstKonsol(0) %>
            <option value="<% = zx %>" <% If CLng(cADD_Kategori) = zx Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
          <% Next %>
        </select>
      </div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Spel</div>  
      <div class="input">
        <input type="hidden" class="fill" id="sSpelID" name="vSpelID" value="<% = cADD_SpelID %>">
        <input type="button" class="browsebutton" value="Välj..." onclick="showPicker('/picker/cms_spel.asp','sSpelID','sSpelText');">
        <input type="button" class="browsebutton" value="Ta bort" style="width: 65px; font-weight: normal;" onclick="clearPicker('sSpelID','sSpelText','Inget spel valt');">
        <input type="text" class="browse fill" style="width: 283px;" id="sSpelText" value="<% = cADD_SpelText %>" disabled>
      </div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Titel</div>
      <div class="input"><input type="text" class="fill notnull" name="vTitel" maxlength="255" value="<% = cADD_Titel %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Betyg</div>
      <div class="input">
        <div class="radiofld" onclick="document.getElementById('vBetyg_0').checked=true;"><input <% If cADD_Betyg = 0 Then Response.Write(" checked") %> type="radio" id="vBetyg_0" name="vBetyg" value=0> Inget</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_1').checked=true;"><input <% If cADD_Betyg = 1 Then Response.Write(" checked") %> type="radio" id="vBetyg_1" name="vBetyg" value=1> 1</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_2').checked=true;"><input <% If cADD_Betyg = 2 Then Response.Write(" checked") %> type="radio" id="vBetyg_2" name="vBetyg" value=2> 2</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_3').checked=true;"><input <% If cADD_Betyg = 3 Then Response.Write(" checked") %> type="radio" id="vBetyg_3" name="vBetyg" value=3> 3</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_4').checked=true;"><input <% If cADD_Betyg = 4 Then Response.Write(" checked") %> type="radio" id="vBetyg_4" name="vBetyg" value=4> 4</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_5').checked=true;"><input <% If cADD_Betyg = 5 Then Response.Write(" checked") %> type="radio" id="vBetyg_5" name="vBetyg" value=5> 5</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_6').checked=true;"><input <% If cADD_Betyg = 6 Then Response.Write(" checked") %> type="radio" id="vBetyg_6" name="vBetyg" value=6> 6</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_7').checked=true;"><input <% If cADD_Betyg = 7 Then Response.Write(" checked") %> type="radio" id="vBetyg_7" name="vBetyg" value=7> 7</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_8').checked=true;"><input <% If cADD_Betyg = 8 Then Response.Write(" checked") %> type="radio" id="vBetyg_8" name="vBetyg" value=8> 8</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_9').checked=true;"><input <% If cADD_Betyg = 9 Then Response.Write(" checked") %> type="radio" id="vBetyg_9" name="vBetyg" value=9> 9</div>
        <div class="radiofld" onclick="document.getElementById('vBetyg_10').checked=true;"><input <% If cADD_Betyg = 10 Then Response.Write(" checked") %> type="radio" id="vBetyg_10" name="vBetyg" value=10> 10</div>
      </div>
    </div>
    
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="texttools">
        <select onchange="if(this.value != 'A'){addText('myText',this.value); this.value='A';}">
          <option value="A"> Rubriker... </option>
          <option class="separator" disabled> &nbsp; </option>
          <option class="levelin" value="h1"> Rubrik 1 </option>
          <option class="levelin" value="h2"> Rubrik 2 </option>
          <option class="levelin" value="h3"> Rubrik 3 </option>
        </select>
        <div class="sep"> </div>
        <img onclick="addText('myText','b');" src="/design/icons/bbcode/b.gif">
        <img onclick="addText('myText','i');" src="/design/icons/bbcode/i.gif">
        <img onclick="addText('myText','u');" src="/design/icons/bbcode/u.gif">
        <div class="sep"> </div>
        <img onclick="addText('myText','url');" src="/design/icons/bbcode/link.gif">
        <div class="sep"> </div>
        <img onclick="addTextEnd('myText','\n[list=1]\n[*]\n[*]\n[*]\n[/list]');" src="/design/icons/bbcode/numlist.gif">
        <img onclick="addTextEnd('myText','\n[list]\n[*]\n[*]\n[*]\n[/list]');" src="/design/icons/bbcode/list.gif">
      </div>
      <textarea id="myText" name="vText"><% = cADD_Text %></textarea>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Teaser</div>  
      <div class="input"><input type="text" class="fill" name="vShort" maxlength="100" value="<% = cADD_Short %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv en kort beskrivande text om recensionen på max 100 tecken.</p> </div>
    
    <div class="in_row">
      <div class="text">Nyckelord</div>  
      <div class="input"><input type="text" class="fill" name="vNyckelord" maxlength="500" value="<% = cADD_Nyckelord %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter recensioner. Seprarera varje nyckelord med ett mellanslag.</p> </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vAnvRec" value="YES" <% If cADD_bAnvRec Then Response.Write(" checked") %>> Användar Rec</div>
      <div class="input">Om en recension är markerad som detta så är den inskickat från sidan, observera att dessa ska bedömmas med mycket lägre krav än vad riktiga recensioner görs.</div>
    </div>
  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  
  <input type="hidden" name="vNotes_cp">
  <input type="hidden" name="vStatus_cp">
  <input type="hidden" name="vPublTyp_cp">
  <input type="hidden" name="vPublDatum_cp">
  <input type="hidden" name="vNySkapare_cp">
  
  </form>
  
  <div class="datablock rect" id="imgholder">
    <div class="legend"><input style="float: right; margin-right: 18px;" type="button" value="Ny bild..." name="saveimg" id="btnew" onclick="show('f_new');mkdisable('btnew', true);" <% If bCantSave Then Response.Write(" disabled") %>> Bilder (<em>Måste sparas separat</em>)</div>
    
    <%
    Set rsIMG = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Bind_Recension_Img LEFT JOIN cms_Bild ON cms_Bind_Recension_Img.brBild = cms_Bild.bID WHERE brSaved = 1 And brRecension = " & CLng(cADD_ID)
    rsIMG.Open SQL, Con, 1, 3
    %>
    
    <!--<div class="textmess<% If Not rsIMG.EOF Then Response.Write("_no") %>" id="f_textmess">Det finns inga bilder.</div>-->
    
    <div class="imgblock" id="f_new" style="display: none;">
      <form id="imgupl_new" method="POST" enctype="multipart/form-data">
        <div class="image"><img src="/design/img_missing.png"></div>
        <div class="fields">
          <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
          <div class="text">Bildtext</div><div class="input"><textarea name="f_brBildText"></textarea></div>
          <input type="hidden" name="f_id" value="0">
          <input type="hidden" name="f_area" value="rec">
          <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
        </div>
        <div class="buttons">
          <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg(0);" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
          <br><br>
          <input style="width: 80px;" type="button" value="Avbryt" name="undoimg" onclick="hide('f_new');mkdisable('btnew', false);" <% If bCantSave Then Response.Write(" disabled") %>>
        </div>
      </form>
    </div>
    
    <%
    Do Until rsIMG.EOF
      %>
      <div class="imgblock" id="f_id<% = rsIMG("brBild") %>">
        <form id="imgupl_id<% = rsIMG("brBild") %>" method="POST" enctype="multipart/form-data">
          <div class="image"><img src="/cms_Img.asp?e=<% = rsIMG("brBild") %>&w=80&h=80"></div>
          <div class="fields">
            <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
            <div class="text">Bildtext<br><br><strong>ID: </strong><% = rsIMG("brBild") %></strong></div><div class="input"><textarea name="f_brBildText"><% = rsIMG("brBildText") %></textarea></div>
            <input type="hidden" name="f_id" value="<% = rsIMG("brBild") %>">
            <input type="hidden" name="f_area" value="rec">
            <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
          </div>
          <div class="buttons">
            <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg(<% = rsIMG("brBild") %>);" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
            <br><br>
            <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg(<% = rsIMG("brBild") %>,<% = cADD_ID %>, 'rec');}" <% If bCantSave Then Response.Write(" disabled") %>>
          </div>
        </form>
      </div>
      <%
      rsIMG.MoveNext
    Loop
    
    rsIMG.Close
    Set rsIMG = Nothing
    %>
  </div>
  
  <!-- ## HIDDENBOX TO COPY ## -->
  <div id="f_hiddenbox" style="display: none;">
    <form id="imgupl_id%ID%" method="POST" enctype="multipart/form-data">
      <div class="image"><img src="%THAIMG%"></div>
      <div class="fields">
        <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
        <div class="text">Bildtext</div><div class="input"><textarea name="f_brBildText">%tetra_f_brBildText%</textarea></div>
        <input type="hidden" name="f_id" value="%ID%">
        <input type="hidden" name="f_area" value="rec">
        <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
      </div>
      <div class="buttons">
        <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg('%ID%');" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
        <br><br>
        <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg('%ID%',<% = cADD_ID %>, 'rec');}" <% If bCantSave Then Response.Write(" disabled") %>>
      </div>
    </form>
  </div>
  <!-- ## HIDDENBOX TO COPY ## -->
  
  <div class="datablock rect morepadding">
    <div class="legend">Noteringar</div>
    <div class="in_row">
      <div class="text">Noteringar</div>
      <div class="input"><textarea id="vNotes" name="vNotes"><% = cADD_Notes %></textarea></div>
    </div>
    <div class="in_text"> <p>Detta fält är för fria noteringar som inte kommer synas för andra än de som kan redigera recensionen. Används även för att förklara varför en recensionen nekades publicering.</p> </div>
  </div>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    <div class="inner" style="text-align: center;">
      <input onclick="cpFlds();saveform('em',0);" name="savebtn" class="save" type="button"value="Spara" <% If bCantSave Then Response.Write(" disabled") %>>
      <input onclick="cpFlds();saveform('em',1);" name="savebtn" class="save_continue" type="button" value="Spara och fortsätt..." <% If bCantSave Then Response.Write(" disabled") %>>
      <input onclick="cpFlds();saveform('em',2);" name="savebtn" class="save_return" type="button" value="Spara och återgå..." <% If bCantSave Then Response.Write(" disabled") %>>
      <input onclick="location.href='_show.asp?<% = sRebuild %>';" name="savebtn" class="cancel" type="button" value="Avbryt">
    </div>
  </div>
  
  <div class="databox info">
    <div class="inner">
      <div id="ajax_loading" style="display: none;">
        <div style="width: 136px; height: 22px; padding: 9px 0 0 38px; background: transparent url('/design/loader.gif') no-repeat; font-weight: bold;" id="do_what">Skickar data...</div>
      </div>
      <div id="ajax_waiting">
        <div style="width: 136px; height: 22px; padding: 9px 0 0 38px; background: transparent url('/design/noload.png') no-repeat; font-weight: bold;" id="dont_do">Avvaktar.</div>
      </div>
      
      <iframe name="processbox" id="processbox" style="width: 174px; height: 180px; display: none;" frameborder=0 src="/_awaiting.asp"></iframe>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Sparad senast</div>
    <div class="inner">
      <%
      If Not rsDB.EOF Then
        saveDate = Trim(CStr(rsDB("rDatumSparad") & " "))
        If saveDate = Empty Then saveDate = "Sparad (datum saknas)" Else saveDate = "Sparad (" & FormatDateTime(saveDate, vbShortDate) & " " & FormatDateTime(saveDate, vbShortTime) & ")"
        isSaved = True
      Else
        saveDate = "Inte sparad"
        isSaved = False
      End If
      %>
      <div class="radio" style="background-image: url('/design/icons/radio_<% If isSaved Then Response.Write("true") Else Response.Write("false") %>.png');" id="savedstatus"><% = saveDate %></div>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Status</div>
    <div class="inner" id="statusarea">
      --
    </div>
  </div>
  
  <script type="text/javascript">getPage("__innerfld.asp?e=<% = cADD_ID %>");</script>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     