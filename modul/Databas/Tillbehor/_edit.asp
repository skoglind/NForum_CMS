<% 
  cON_PAGE = "Hantera tillbehör - Tillbehör - CMS"
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
    SQL = "SELECT * FROM cms_Tillbehor WHERE iID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Synlig             = True
      cADD_Has_SinglePlay     = True
      
      cADD_Konsol             = sKonsol
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("iID")
      cADD_Konsol             = rsDB("iKonsol")
      cADD_Titel              = rsDB("iStandard_Titel")
      cADD_Text               = sEncode(rsDB("iTextM"))
      cADD_Nyckelord          = sEncode(rsDB("iNyckelord"))
      
      cADD_Synlig             = rsDB("iSynlig")
      
      cADD_StandardTitel      = rsDB("iStandard_Titel")
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
      cpVal('vSynlig');
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Tillbehör</div>
    
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
    <div class="legend" style="margin-bottom: 5px;"><input style="float: right;" type="button" value="Ny titel..." onclick="addMe('titles',getSlumpID(),'','','','',0,'','');"> Titlar</div>

  </div>
    
  <div class="datablock rect morepadding">
    <div class="legend">Övriga uppgifter</div>
    
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
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter tillbehör. Seprarera varje nyckelord med ett mellanslag.</p> </div>
    
  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
  <input type="hidden" name="q" value="<% = sQ %>">
  
  <input type="hidden" name="vSynlig_cp">
  
  </form>
  
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
        <input class="fill notnull" type="text" id="vTitel_mmID" name="vTitel_mmID" maxlength="255" style="width: 252px;">
      </div>
      
      <div class="text">Release</div>
      <div class="input" style="width: 120px;">
        <input class="fill" type="text" id="vRelease_mmID" name="vRelease_mmID" maxlength="40" style="width: 102px;">
      </div>
      
      <div class="text" style="width: 70px;">Regionskod</div>
      <div class="input" style="width: 224px;">
        <input class="fill" type="text" id="vRegionskod_mmID" name="vRegionskod_mmID" maxlength="50" style="width: 214px;">
      </div>
      
      <div class="boxart_holder">
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_1_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_1_mmID','mmID',1,'Box - Framsida',<% = cADD_ID %>)" id="upload_valj_1_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',1,'addon');}" name="uplbtn">
          </div>
          <div class="titel">Box - Framsida</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_2_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_2_mmID','mmID',2,'Box - Baksida',<% = cADD_ID %>)" id="upload_valj_2_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',2,'addon');}" name="uplbtn">
          </div>
          <div class="titel">Box - Baksida</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_3_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_3_mmID','mmID',3,'Manual',<% = cADD_ID %>)" id="upload_valj_3_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',3,'addon');}" name="uplbtn">
          </div>
          <div class="titel">Manual</div>
        </div>
        
        <div class="boxart">
          <div class="pic"><img src="/design/noimg.gif" id="upload_pic_4_mmID" onclick="showPic(this.src);" style="cursor: pointer;"></div>
          <div class="controls">
            <input type="button" value="Välj..." style="font-weight: bold;" onclick="show_uploadbox('upload_valj_4_mmID','mmID',4,'Tillbehör',<% = cADD_ID %>)" id="upload_valj_4_mmID">
            <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort bilden?')){boxdeleteimg('mmID',4,'addon');}" name="uplbtn">
          </div>
          <div class="titel">Tillbehör</div>
        </div>
      </div>
      
      <input type="hidden" id="vExtra_mmID">
      <input type="hidden" id="vUtgivareID_mmID">
      <input type="hidden" id="vUtgivareText_mmID">
      
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
      SQL = "SELECT * FROM cms_Tillbehortitlar WHERE tTillbehorID = " & CLng(listID) & " ORDER BY tID ASC"
      rsT.Open SQL, con
      
      If rsT.EOF Then
        ' Ne
      End If
    
      Do Until rsT.EOF
        cADD_Titel        = rsT("tTitel")
        cADD_Release      = rsT("tRelease")
        cADD_Region       = rsT("tRegion")
        cADD_SortNo       = rsT("tSortNo")
        cADD_Regionskod   = rsT("tRegionskod")
        
        cADD_ID           = rsT("tID")
        
        cADD_BoxFram_no   = rsT("tBoxart_BoxFram")
        cADD_BoxBak_no    = rsT("tBoxart_BoxBak")
        cADD_Manual_no    = rsT("tBoxart_Manual")
        cADD_Kassett_no   = rsT("tBoxart_Tillbehor")
        
        If CLng(cADD_BoxFram_no) > 0 Then cADD_BoxFram = "/cms_Img.asp?e=" & cADD_BoxFram_no & "&w=80&h=80" Else cADD_BoxFram = "/design/noimg.gif"
        If CLng(cADD_BoxBak_no) > 0  Then cADD_BoxBak = "/cms_Img.asp?e=" & cADD_BoxBak_no & "&w=80&h=80"   Else cADD_BoxBak  = "/design/noimg.gif"
        If CLng(cADD_Manual_no) > 0  Then cADD_Manual = "/cms_Img.asp?e=" & cADD_Manual_no & "&w=80&h=80"   Else cADD_Manual  = "/design/noimg.gif"
        If CLng(cADD_Kassett_no) > 0 Then cADD_Kassett = "/cms_Img.asp?e=" & cADD_Kassett_no & "&w=80&h=80" Else cADD_Kassett = "/design/noimg.gif"
               
        Response.Write("addMe('titles'," & cADD_ID & ",'" & jEncode(cADD_Titel) & "','','" & jEncode(cADD_Release) & "','" & jEncode(cADD_Regionskod) & "',0,''," & cADD_Region & ");" & vbCrlf)
        Response.Write("selectValue('vRegion_" & cADD_ID & "'," & cADD_Region & ");" & vbCrlf)
        Response.Write("selectValue('vSortNo_" & cADD_ID & "'," & cADD_SortNo & ");" & vbCrlf)
        Response.Write("setBoxart(" & cADD_ID & ",'" & cADD_BoxFram & "','" & cADD_BoxBak & "','" & cADD_Manual & "','" & cADD_Kassett & "');" & vbCrlf)
        
        rsT.MoveNext
      Loop
      
      rsT.Close
      Set rsT = Nothing
      %>
      
      <% If lPBStatus = "NewPost" Then %>addMe('titles',getSlumpID(),'','','','',0,'','');<% End If %>
      setRadio("vStandardTitel","<% = cADD_StandardTitel %>");
  </script>
  
  <div id="uploadbox">
    <div class="holder">
      <div class="empty"></div>
      <div class="data">
        <form method="POST" enctype="multipart/form-data" id="boxart_upl">
          <input type="hidden" name="uID" id="upload_ID">
          <input type="hidden" name="uArt" id="upload_Art">
          <input type="hidden" name="uArea" value="addon">
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
        saveDate = Trim(CStr(rsDB("iDatumSparad") & " "))
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