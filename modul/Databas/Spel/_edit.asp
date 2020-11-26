<% 
  cON_PAGE = "Hantera spel - Spel - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS4") Then Response.Redirect("/")
  %>

  <%
  lID = Request.QueryString("e")
  If Not IsNumeric(lID) Or lID = Empty Then lID = 0
  lID = CLng(lID)
  
  sKonsol = noFnutt(Request.QueryString("f"))
  If Not IsNumeric(sKonsol) Or sKonsol = Empty Then sKonsol = 0
  sKonsol = CLng(sKonsol)
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM cms_Spel LEFT JOIN cms_Foretag ON cms_Spel.sUtvecklare = cms_Foretag.fID WHERE sID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Synlig             = True
      cADD_Has_SinglePlay     = True
      
      cADD_ESRB               = 0
      cADD_PEGI               = 0
      cADD_Spelare            = 0
      
      cADD_Utvecklare         = 0
      cADD_UtvecklareNamn     = "Ingen utvecklare vald"
      
      cADD_Konsol             = sKonsol
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("sID")
      cADD_Konsol             = rsDB("sKonsol")
      cADD_Titel              = rsDB("sStandard_Titel")
      cADD_Text               = sEncode(rsDB("sTextM"))
      cADD_Nyckelord          = sEncode(rsDB("sNyckelord"))
      
      If IsNull(rsDB("fID")) Then
        cADD_Utvecklare         = 0
        cADD_UtvecklareNamn     = "Ingen utvecklare vald"
      Else
        cADD_Utvecklare         = CLng(rsDB("sUtvecklare"))
        cADD_UtvecklareNamn     = sEncode(rsDB("fNamn"))
      End If
      
      cADD_Has_SinglePlay     = rsDB("sSingleplayer")
      cADD_Has_MultiPlay      = rsDB("sMultiplayer")
      cADD_Has_Online         = rsDB("sOnline")
      cADD_License            = rsDB("sOLicensierad")
      
      cADD_Synlig             = rsDB("sSynlig")
      
      cADD_ESRB               = CLng(rsDB("sESRB"))
      cADD_PEGI               = CLng(rsDB("sPEGI"))
      cADD_Spelare            = CLng(rsDB("sAntalSpelare"))
      
      cADD_StandardTitel      = rsDB("sStandard_Titel")
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sQ = Trim(Left(MakeLegal_Large(Request.QueryString("q")), 255))
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  Call GetAlfa(Request.QueryString("alfa"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa & "&k=" & sKonsol & "&q=" & sQ
  ' ##################
  %>
  
  <script type="text/javascript">
    function cpFlds() {
      cpVal('vESRB');
      cpVal('vPEGI');
      cpVal('vAntalSpelare');
      cpVal('vSynlig');
      cpVal('vSinglePlay');
      cpVal('vMultiPlay');
      cpVal('vOnline');
      cpVal('vLicense');
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Konsol</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Konsol</div>
      <div class="input">
        <select name="vKonsol" id="vKonsol">
          <% For zx = 1 To lstKonsol(0) %>
            <option value="<% = zx %>" <% If CLng(cADD_Konsol) = zx Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
          <% Next %>
        </select>
      </div>
    </div>
  </div>
    
  <div class="datablock rect morepadding" id="titles">
    <div class="legend" style="margin-bottom: 5px;"><input style="float: right;" type="button" value="Ny titel..." onclick="addMe('titles',getSlumpID(),'','','','',0,'Ingen utgivare vald','');"> Titlar</div>

  </div>
    
  <div class="datablock rect morepadding">
    <div class="legend">Övriga uppgifter</div>
  
    <div class="in_row">
      <div class="text">Utvecklare</div>  
      <div class="input">
        <input type="hidden" class="fill" id="sUtvecklareID" name="vUtvecklare" value="<% = cADD_Utvecklare %>">
        <input type="button" class="browsebutton" value="Välj..." onclick="showPicker('/picker/cms_Foretag.asp','sUtvecklareID','sUtvecklareText');">
        <input type="button" class="browsebutton" value="Ta bort" style="width: 65px; font-weight: normal;" onclick="clearPicker('sUtvecklareID','sUtvecklareText','Ingen utvecklare vald');">
        <input type="text" class="browse fill" style="width: 283px;" id="sUtvecklareText" value="<% = cADD_UtvecklareNamn %>" disabled>
      </div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="twopart">Genre:</div>
      <div class="twopart">Spelgrupp:</div>
    </div>
    
    <div class="in_row">
      <% Set rsList = Server.CreateObject("ADODB.Recordset") %>
    
      <% 
        SQL = "SELECT * FROM cms_Spelgengre LEFT JOIN cms_Bind_Spel_Genre ON cms_Spelgengre.gID = cms_Bind_Spel_Genre.bgGenre AND cms_Bind_Spel_Genre.bgSpel = " & CLng(cADD_ID) & " ORDER BY gNamn ASC"
        rsList.Open SQL, con
      %>
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="genre_radio" value="0" id="genre" onclick="getToggle('genre');"></div> <div class="lbl" onclick="toggletrue('genre');getToggle('genre');">Ingen genre</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="genre" value="<% = rsList("gID") %>" id="genre<% = rsList("gID") %>" onclick="setToggle('genre');" <% If Not IsNull(rsList("bgSpel")) Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('genre<% = rsList("gID") %>');setToggle('genre');"><% = rsList("gNamn") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
        SQL = "SELECT * FROM cms_Spelserier LEFT JOIN cms_Bind_Spel_Spelserie ON cms_Spelserier.ssID = cms_Bind_Spel_Spelserie.bsSpelSerie AND cms_Bind_Spel_Spelserie.bsSpel = " & CLng(cADD_ID) & " ORDER BY ssNamn ASC"
        rsList.Open SQL, con
      %>
      
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="grupp_radio" value="0" id="grupp" onclick="getToggle('grupp');"></div> <div class="lbl" onclick="toggletrue('grupp');getToggle('grupp');">Ingen spelgrupp</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="grupp" value="<% = rsList("ssID") %>" id="grupp<% = rsList("ssID") %>" onclick="setToggle('grupp');" <% If Not IsNull(rsList("bsSpel")) Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('grupp<% = rsList("ssID") %>');setToggle('grupp');">[<% If rsList("ssSerien") Then Response.Write("S") Else Response.Write("G") %>] <% = rsList("ssNamn") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
      %>
      
      <% Set rsList = Nothing %>
      
      <script type="text/javascript">
        setToggle('genre');
        setToggle('grupp');
      </script>
    </div>
    
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
      <div class="text">Nyckelord</div>  
      <div class="input"><input type="text" class="fill" name="vNyckelord" maxlength="500" value="<% = cADD_Nyckelord %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter spel. Seprarera varje nyckelord med ett mellanslag.</p> </div>
    
  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
  <input type="hidden" name="q" value="<% = sQ %>">
  
  <input type="hidden" name="vESRB_cp">
  <input type="hidden" name="vPEGI_cp">
  <input type="hidden" name="vAntalSpelare_cp">
  <input type="hidden" name="vSynlig_cp">
  <input type="hidden" name="vSinglePlay_cp">
  <input type="hidden" name="vMultiPlay_cp">
  <input type="hidden" name="vOnline_cp">
  <input type="hidden" name="vLicense_cp">
  
  </form>
  
  <div class="datablock rect" id="imgholder">
    <div class="legend"><input style="float: right;" type="button" value="Ny bild..." name="saveimg" id="btnew" onclick="show('f_new');mkdisable('btnew', true);" <% If bCantSave Then Response.Write(" disabled") %>> Bilder (<em>Måste sparas separat</em>)</div>
    
    <%
    Set rsIMG = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Bind_Spel_Img LEFT JOIN cms_Bild ON cms_Bind_Spel_Img.bsBild = cms_Bild.bID WHERE bsSaved = 1 And bsSpel = " & CLng(cADD_ID)
    rsIMG.Open SQL, Con, 1, 3
    %>
    
    <!--<div class="textmess<% If Not rsIMG.EOF Then Response.Write("_no") %>" id="f_textmess">Det finns inga bilder.</div>-->
    
    <div class="imgblock" id="f_new" style="display: none;">
      <form id="imgupl_new" method="POST" enctype="multipart/form-data">
        <div class="image"><img src="/design/img_missing.png"></div>
        <div class="fields">
          <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
          <div class="text">Bildtext</div><div class="input"><textarea name="f_bsBildText"></textarea></div>
          <input type="hidden" name="f_id" value="0">
          <input type="hidden" name="f_area" value="game">
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
      <div class="imgblock" id="f_id<% = rsIMG("bsBild") %>">
        <form id="imgupl_id<% = rsIMG("bsBild") %>" method="POST" enctype="multipart/form-data">
          <div class="image"><img src="/cms_Img.asp?e=<% = rsIMG("bsBild") %>&w=80&h=80"></div>
          <div class="fields">
            <div class="text">Bild</div><div class="input"><input name="f_file" type="file" size="43"></div>
            <div class="text">Bildtext</div><div class="input"><textarea name="f_bsBildText"><% = rsIMG("bsBildText") %></textarea></div>
            <input type="hidden" name="f_id" value="<% = rsIMG("bsBild") %>">
            <input type="hidden" name="f_area" value="game">
            <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
          </div>
          <div class="buttons">
            <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg(<% = rsIMG("bsBild") %>);" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
            <br><br>
            <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg(<% = rsIMG("bsBild") %>,<% = cADD_ID %>, 'game');}" <% If bCantSave Then Response.Write(" disabled") %>>
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
        <div class="text">Bildtext</div><div class="input"><textarea name="f_bsBildText">%tetra_f_bsBildText%</textarea></div>
        <input type="hidden" name="f_id" value="%ID%">
        <input type="hidden" name="f_area" value="game">
        <input type="hidden" name="f_objid" value="<% = cADD_ID %>">
      </div>
      <div class="buttons">
        <input style="width: 80px; font-weight: bold;" type="button" value="Spara" onclick="uploadimg('%ID%');" name="saveimg" <% If bCantSave Then Response.Write(" disabled") %>>
        <br><br>
        <input style="width: 80px;" type="button" value="Radera" name="undoimg" onclick="if(confirm('Vill du radera bilden?')){deleteimg('%ID%',<% = cADD_ID %>, 'game');}" <% If bCantSave Then Response.Write(" disabled") %>>
      </div>
    </form>
  </div>
  <!-- ## HIDDENBOX TO COPY ## -->
  
  <!-- ## HIDDENBOX TO COPY ## -->
  <div id="t_hiddenbox" style="display: none;">
    <div class="legend">
      <div class="iRadio"><input type="radio" style="margin: 0;" name="vStandardTitel" value="mmID"></div>
      <div class="iText">Standard</div> 
      <div class="iDelete"><input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort titeln?')){removeMe('titles','TitelRow_mmID');}"></div>
    </div>
  
    <div class="frame">
      <div class="text">Titel</div>
      <div class="input" style="width: 430px;">
        <select style="width: 112px;" id="vRegion_mmID" name="vRegion_mmID">
          <% Call ListRegion(0) %>
        </select>
        <select style="width: 48px;" id="vSortNo_mmID" name="vSortno_mmID">
          <option value=0> -- </option>
          <% For zx = 1 To 25 %>
            <option value=<% = zx %>> <% = zx %> </option>
          <% Next %>
        </select>
        <input class="fill notnull" type="text" id="vTitel_mmID" name="vTitel_mmID" maxlength="255" style="width: 252px;" onkeyup="reduceLoad('vTitel_mmID', document.getElementById('vKonsol').value, 'mmID', <% = cADD_ID %>);" onblur="hideTitleExists();">
      </div>
      
      <div class="text">Extra information</div>
      <div class="input" style="width: 430px;">
        <input class="fill" type="text" id="vExtra_mmID" name="vExtra_mmID" maxlength="255" style="width: 420px; color: #888;">
      </div>
      
      <div class="text">Release</div>
      <div class="input" style="width: 120px;">
        <input class="fill" type="text" id="vRelease_mmID" name="vRelease_mmID" maxlength="40" style="width: 102px;">
      </div>
      
      <div class="text" style="width: 70px;">Regionskod</div>
      <div class="input" style="width: 224px;">
        <input class="fill" type="text" id="vRegionskod_mmID" name="vRegionskod_mmID" maxlength="50" style="width: 214px;">
      </div>
      
      <div class="text">Utgivare</div>
      <div class="input" style="width: 430px;">
        <input type="hidden" class="fill" id="vUtgivareID_mmID" name="vUtgivareID_mmID">
        <input type="button" class="browsebutton" value="Välj..." onclick="showPicker('/picker/cms_Foretag.asp','vUtgivareID_mmID','vUtgivareText_mmID');">
        <input type="button" class="browsebutton" value="Ta bort" style="width: 65px; font-weight: normal;" onclick="clearPicker('vUtgivareID_mmID','vUtgivareText_mmID','Ingen utgivare vald');">
        <input type="text" class="browse fill" style="width: 275px;" id="vUtgivareText_mmID" name="vUtgivareText_mmID" disabled>
      </div>
      
      <div class="boxart_holder">
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_1_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_1_mmID','mmID',1,'Box - Framsida',<% = cADD_ID %>)" id="upload_valj_1_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',1,'game');}" name="uplbtn">
          </div>
          <div class="titel">Box - Framsida</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_2_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_2_mmID','mmID',2,'Box - Baksida',<% = cADD_ID %>)" id="upload_valj_2_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',2,'game');}" name="uplbtn">
          </div>
          <div class="titel">Box - Baksida</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_3_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_3_mmID','mmID',3,'Manual',<% = cADD_ID %>)" id="upload_valj_3_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',3,'game');}" name="uplbtn">
          </div>
          <div class="titel">Manual</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_4_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_4_mmID','mmID',4,'Kassett',<% = cADD_ID %>)" id="upload_valj_4_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',4,'game');}" name="uplbtn">
          </div>
          <div class="titel">Kassett</div>
        </div>
      </div>
      
    </div>
  </div>
  <!-- ## HIDDENBOX TO COPY ## -->
  
  <script type="text/javascript">
    <%
      If CLng(cADD_ID) = 0 Then 
        listID = -1
      Else
        listID = CLng(cADD_ID)
      End If
    
      Set rsT = Server.CreateObject("ADODB.Recordset") 
      SQL = "SELECT * FROM cms_Speltitlar LEFT JOIN cms_Foretag ON cms_Speltitlar.tUtgivare = cms_Foretag.fID WHERE tSpelID = " & CLng(listID) & " ORDER BY tID ASC"
      rsT.Open SQL, con
      
      If rsT.EOF Then
        '
      End If
    
      Do Until rsT.EOF
        If IsNull(rsT("fID")) Then
          cADD_Utgivare           = 0
          cADD_UtgivareNamn       = "Ingen utgivare vald"
        Else
          cADD_Utgivare           = CLng(rsT("tUtgivare"))
          cADD_UtgivareNamn       = sEncode(rsT("fNamn"))
        End If
        
        cADD_Titel        = rsT("tTitel")
        cADD_Extra        = rsT("tExtra")
        cADD_Release      = rsT("tRelease")
        cADD_Region       = rsT("tRegion")
        cADD_SortNo       = rsT("tSortNo")
        cADD_Regionskod   = rsT("tRegionskod")
        
        cADD_ID           = rsT("tID")
        
        cADD_BoxFram_no   = rsT("tBoxart_BoxFram")
        cADD_BoxBak_no    = rsT("tBoxart_BoxBak")
        cADD_Manual_no    = rsT("tBoxart_Manual")
        cADD_Kassett_no   = rsT("tBoxart_Kassett")
        
        If CLng(cADD_BoxFram_no) > 0 Then cADD_BoxFram = "/cms_Img.asp?e=" & cADD_BoxFram_no & "&w=80&h=80" Else cADD_BoxFram = "/design/noimg.gif"
        If CLng(cADD_BoxBak_no) > 0  Then cADD_BoxBak = "/cms_Img.asp?e=" & cADD_BoxBak_no & "&w=80&h=80"   Else cADD_BoxBak  = "/design/noimg.gif"
        If CLng(cADD_Manual_no) > 0  Then cADD_Manual = "/cms_Img.asp?e=" & cADD_Manual_no & "&w=80&h=80"   Else cADD_Manual  = "/design/noimg.gif"
        If CLng(cADD_Kassett_no) > 0 Then cADD_Kassett = "/cms_Img.asp?e=" & cADD_Kassett_no & "&w=80&h=80" Else cADD_Kassett = "/design/noimg.gif"
               
        Response.Write("addMe('titles'," & cADD_ID & ",'" & jEncode(cADD_Titel) & "','" & jEncode(cADD_Extra) & "','" & jEncode(cADD_Release) & "','" & jEncode(cADD_Regionskod) & "'," & cADD_Utgivare & ",'" & jEncode(cADD_UtgivareNamn) & "'," & cADD_Region & ");" & vbCrlf)
        Response.Write("selectValue('vRegion_" & cADD_ID & "'," & cADD_Region & ");" & vbCrlf)
        Response.Write("selectValue('vSortNo_" & cADD_ID & "'," & cADD_SortNo & ");" & vbCrlf)
        Response.Write("setBoxart(" & cADD_ID & ",'" & cADD_BoxFram & "','" & cADD_BoxBak & "','" & cADD_Manual & "','" & cADD_Kassett & "');" & vbCrlf)
        
        rsT.MoveNext
      Loop
      
      rsT.Close
      Set rsT = Nothing
      %>
      
      <% If lPBStatus = "NewPost" Then %>addMe('titles',getSlumpID(),'','','','',0,'Ingen utgivare vald','');<% End If %>
      setRadio("vStandardTitel","<% = cADD_StandardTitel %>");
  </script>
  
  <div id="uploadbox">
    <div class="holder">
      <div class="empty"></div>
      <div class="data">
        <form method="POST" enctype="multipart/form-data" id="boxart_upl">
          <input type="hidden" name="uID" id="upload_ID">
          <input type="hidden" name="uArt" id="upload_Art">
          <input type="hidden" name="uArea" value="game">
          <input type="hidden" name="uGameID" id="game_ID">
          
          <input type="file" size=23 id="upload_File" name="upload_Boxart">
          <input class="btn" type="button" value="Ladda upp" style="font-weight: bold;" onclick="boxuploadimg();hide_uploadbox();" name="uplbtn">
          <input class="btn" type="button" value="Avbryt" onclick="hide_uploadbox();">
        </form>
      </div>
      <div class="titel" id="upload_Text">Box - Framsida</div>
    </div>    
  </div>
  
  <div id="titlechecker">  
  </div>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    <div class="inner" style="text-align: center;">
      <input onclick="cpFlds();saveform('em',0);" name="savebtn" class="save" type="button" value="Spara" <% If bCantSave Then Response.Write(" disabled") %>>
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
        saveDate = Trim(CStr(rsDB("sDatumSparad") & " "))
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
    <div class="label">Visningsalternativ</div>
    <div class="inner">
      <div class="chkbox"><input type="checkbox" value="YES" id="vSynlig" name="vSynlig" <% If cADD_Synlig Then Response.Write(" checked") %>></div>
      <div class="chkbox_text">Synlig (syns på N-Forum)</div>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Alternativ</div>
    <div class="inner">
      <div class="field">
        <select id="vESRB" name="vESRB">
          <option value="A"> ESRB Rating... </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="0" class="levelin" <% If cADD_ESRB = 0 Then Response.Write(" selected") %>> ESRB | Ej angiven </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="1" class="levelin" <% If cADD_ESRB = 1 Then Response.Write(" selected") %>> ESRB | EC (Early Childhood) </option>
          <option value="2" class="levelin" <% If cADD_ESRB = 2 Then Response.Write(" selected") %>> ESRB | E (Everyone)  </option>
          <option value="3" class="levelin" <% If cADD_ESRB = 3 Then Response.Write(" selected") %>> ESRB | E10+ (Everyone 10 and older) </option>
          <option value="4" class="levelin" <% If cADD_ESRB = 4 Then Response.Write(" selected") %>> ESRB | T (Teen) </option>
          <option value="5" class="levelin" <% If cADD_ESRB = 5 Then Response.Write(" selected") %>> ESRB | M (Mature) </option>
          <option value="6" class="levelin" <% If cADD_ESRB = 6 Then Response.Write(" selected") %>> ESRB | AO (Adults Only) </option>
        </select>
      </div>
      
      <div class="field">
        <select id="vPEGI" name="vPEGI">
          <option value="A"> PEGI Rating... </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="0" class="levelin" <% If cADD_PEGI = 0 Then Response.Write(" selected") %>> PEGI | Ej angiven </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="1" class="levelin" <% If cADD_PEGI = 1 Then Response.Write(" selected") %>> PEGI | 3+ </option>
          <option value="2" class="levelin" <% If cADD_PEGI = 2 Then Response.Write(" selected") %>> PEGI | 7+  </option>
          <option value="3" class="levelin" <% If cADD_PEGI = 3 Then Response.Write(" selected") %>> PEGI | 12+ </option>
          <option value="4" class="levelin" <% If cADD_PEGI = 4 Then Response.Write(" selected") %>> PEGI | 16+ </option>
          <option value="5" class="levelin" <% If cADD_PEGI = 5 Then Response.Write(" selected") %>> PEGI | 18+ </option>
        </select>
      </div>   
     
      <div class="innerseparator"> </div>
      
      <div class="chkbox"><input type="checkbox" value="YES" id="vSinglePlay" name="vSinglePlay" <% If cADD_Has_SinglePlay Then Response.Write(" checked") %>></div>
      <div class="chkbox_text">Singleplayer</div>
      
      <div class="chkbox"><input type="checkbox" value="YES" id="vMultiPlay" name="vMultiPlay" <% If cADD_Has_MultiPlay Then Response.Write(" checked") %>></div>
      <div class="chkbox_text">Multiplayer</div>
      
      <div class="field">
        <select id="vAntalSpelare" name="vAntalSpelare">
          <option value="A"> Antal spelare... </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="0" class="levelin" <% If cADD_Spelare = 0 Then Response.Write(" selected") %>> Spelare | Ej angivet </option>
          <option class="separator" disabled> &nbsp; </option>
          <option value="1" class="levelin" <% If cADD_Spelare = 1 Then Response.Write(" selected") %>> Spelare | 1 st </option>
          <option value="2" class="levelin" <% If cADD_Spelare = 2 Then Response.Write(" selected") %>> Spelare | 2 st  </option>
          <option value="3" class="levelin" <% If cADD_Spelare = 3 Then Response.Write(" selected") %>> Spelare | 3 st </option>
          <option value="4" class="levelin" <% If cADD_Spelare = 4 Then Response.Write(" selected") %>> Spelare | 4 st </option>
          <option value="5" class="levelin" <% If cADD_Spelare = 5 Then Response.Write(" selected") %>> Spelare | Fler än 4 st </option>
        </select>
      </div>  
      
      <div class="innerseparator"> </div>
      
      <div class="chkbox"><input type="checkbox" value="YES" id="vOnline" name="vOnline" <% If cADD_Has_Online Then Response.Write(" checked") %>></div>
      <div class="chkbox_text">Online (spelläge)</div>
      
      <div class="innerseparator"> </div>
      
      <div class="chkbox"><input type="checkbox" value="YES" id="vLicense" name="vLicense" <% If cADD_License Then Response.Write(" checked") %>></div>
      <div class="chkbox_text">Olicensierad</div>
    </div>
  </div>
  
  <div class="databox info" style="display: none;">
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