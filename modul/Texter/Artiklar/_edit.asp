<% 
  cON_PAGE = "Hantera artikel - Artiklar - CMS"
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
  If NOT GetAcc("CMS111") Then sBFilter = " AND aaSkapadAv = " & cCMS_ID & " AND NOT aaStatus = 0"
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT *, anvDB1.aNamn AS Anv1, anvDB2.aNamn AS Anv2 " & _
          "FROM (cms_Artiklar " & _ 
          "LEFT JOIN fsBB_Anv AS AnvDB1 ON cms_Artiklar.aaSkapadAv = AnvDB1.aID) " & _
          "LEFT JOIN fsBB_Anv AS AnvDB2 ON cms_Artiklar.aaPubliceradAv = AnvDB2.aID " & _
          "WHERE aaID = " & CLng(lID) & sBFilter
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Kategori           = 1
      cADD_Betyg              = 0
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("aaID")
      cADD_Kategori           = rsDB("aaKategori")
      cADD_Titel              = sEncode(rsDB("aaTitel"))
      cADD_Text               = sEncode(rsDB("aaText"))
      cADD_Notes              = sEncode(rsDB("aaNotes"))
      cADD_Short              = sEncode(rsDB("aaShort"))
      cADD_Nyckelord          = sEncode(rsDB("aaNyckelord"))
      
      cADD_bAnvArt            = rsDB("aaAnvandarArt")
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida
  ' ##################
  
  If Not rsDB.EOF Then If Not GetAcc("CMS11") And rsDB("aaStatus") <> 1 Then bCantSave = True
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
    <div class="legend">Hantera artikel</div>
    
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
      <div class="text">Titel</div>
      <div class="input"><input type="text" class="fill notnull" name="vTitel" maxlength="255" value="<% = cADD_Titel %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="texttools">
        <select onchange="if(this.value != 'A'){addText('myText',this.value); this.value='A';}">
          <option value="A"> Rubriker... </option>
          <option class="separator" disabled> &nbsp; </option>
          <option class="levelin" value="rubrik"> Rubrik </option>
        </select>
        <div class="sep"> </div>
        <img onclick="addText('myText','b');" src="/design/icons/bbcode/b.gif">
        <img onclick="addText('myText','i');" src="/design/icons/bbcode/i.gif">
        <img onclick="addText('myText','u');" src="/design/icons/bbcode/u.gif">
        <div class="sep"> </div>
        <img onclick="addText('myText','url');" src="/design/icons/bbcode/link.gif">
        <div class="sep"> </div>
        <img onclick="addTextEnd('myText','\n[list]\n[*]\n[*]\n[*]\n[/list]');" src="/design/icons/bbcode/list.gif">
      </div>
      <textarea id="myText" name="vText" style="height: 500px;"><% = cADD_Text %></textarea>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Teaser</div>  
      <div class="input"><input type="text" class="fill" name="vShort" maxlength="100" value="<% = cADD_Short %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv en kort beskrivande text om artikeln på max 100 tecken.</p> </div>
    
    <div class="in_row">
      <div class="text">Nyckelord</div>  
      <div class="input"><input type="text" class="fill" name="vNyckelord" maxlength="500" value="<% = cADD_Nyckelord %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter artiklar. Seprarera varje nyckelord med ett mellanslag.</p> </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vAnvArt" value="YES" <% If cADD_bAnvArt Then Response.Write(" checked") %>> Användar Art</div>
      <div class="input">Om en artikel är markerad som detta så är den inskickat från sidan, observera att dessa ska bedömmas med mycket lägre krav än vad riktiga artiklar görs.</div>
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
    SQL = "SELECT * FROM cms_Bind_Artikel_Img LEFT JOIN cms_Bild ON cms_Bind_Artikel_Img.baBild = cms_Bild.bID WHERE baSaved = 1 And baArtikel = " & CLng(cADD_ID)
    rsIMG.Open SQL, Con, 1, 3
    %>
    
    <!--<div class="textmess<% If Not rsIMG.EOF Then Response.Write("_no") %>" id="f_textmess">Det finns inga bilder.</div>-->
    
    <div class="imgblock" id="f_new" style="display: none;">
      <form id="imgupl_new" method="POST" enctype="multipart/form-data">
        <div class="image"><img src="/design/img_missing.png"></div>
        <div class="fields">
          <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
          <div class="text">Bildtext</div><div class="input"><textarea name="f_baBildText"></textarea></div>
          <input type="hidden" name="f_id" value="0">
          <input type="hidden" name="f_area" value="art">
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
      <div class="imgblock" id="f_id<% = rsIMG("baBild") %>">
        <form id="imgupl_id<% = rsIMG("baBild") %>" method="POST" enctype="multipart/form-data">
          <div class="image"><img src="/cms_Img.asp?e=<% = rsIMG("baBild") %>&w=80&h=80"></div>
          <div class="fields">
            <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
            <div class="text">Bildtext<br><br><strong>ID: </strong><% = rsIMG("baBild") %></strong></div><div class="input"><textarea name="f_baBildText"><% = rsIMG("baBildText") %></textarea></div>
            <input type="hidden" name="f_id" value="<% = rsIMG("baBild") %>">
            <input type="hidden" name="f_area" value="art">
            <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
          </div>
          <div class="buttons">
            <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg(<% = rsIMG("baBild") %>);" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
            <br><br>
            <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg(<% = rsIMG("baBild") %>,<% = cADD_ID %>, 'art');}" <% If bCantSave Then Response.Write(" disabled") %>>
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
        <div class="text">Bildtext</div><div class="input"><textarea name="f_baBildText">%tetra_f_baBildText%</textarea></div>
        <input type="hidden" name="f_id" value="%ID%">
        <input type="hidden" name="f_area" value="art">
        <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
      </div>
      <div class="buttons">
        <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg('%ID%');" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
        <br><br>
        <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg('%ID%',<% = cADD_ID %>, 'art');}" <% If bCantSave Then Response.Write(" disabled") %>>
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
    <div class="in_text"> <p>Detta fält är för fria noteringar som inte kommer synas för andra än de som kan redigera artikeln. Används även för att förklara varför en artikel nekades publicering.</p> </div>
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
        saveDate = Trim(CStr(rsDB("aaDatumSparad") & " "))
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