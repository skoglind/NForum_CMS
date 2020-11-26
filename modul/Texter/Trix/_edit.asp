<% 
  cON_PAGE = "Hantera tips & trix - Tips & Trix (Spel) - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS111") Then Response.Redirect("/")
  %>

  <%
  lID = Request.QueryString("e")
  If Not IsNumeric(lID) Or lID = Empty Then lID = 0
  lID = CLng(lID)
  
  ' #### BEHÖRIGHET ####
  'If NOT GetAcc("CMS111") Then sBFilter = " AND rSkapadAv = " & cCMS_ID & " AND NOT rStatus = 0"
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT *, anvDB1.aNamn AS Anv1, anvDB2.aNamn AS Anv2 " & _
          "FROM (cms_Speltrix " & _ 
          "LEFT JOIN cms_Spel ON sID = xSpelID " & _
          "LEFT JOIN cms_SpelTitlar ON tID = cms_Spel.sStandard_Titel " & _
          "LEFT JOIN fsBB_Anv AS AnvDB1 ON cms_Speltrix.xSkapadAv = AnvDB1.aID) " & _
          "LEFT JOIN fsBB_Anv AS AnvDB2 ON cms_Speltrix.xPubliceradAv = AnvDB2.aID " & _
          "WHERE xID = " & CLng(lID) & sBFilter
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_SpelID             = 0
      cADD_SpelText           = "Inget spel valt"
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("xID")
      cADD_Titel              = sEncode(rsDB("xTitel"))
      cADD_Text               = sEncode(rsDB("xTextM"))
      cADD_SpelID             = rsDB("xSpelID")
      cADD_SpelText           = lstKonsolXShort(rsDB("sKonsol")) & " | " & rsDB("tTitel")
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida
  ' ##################

  %>
  
  <script type="text/javascript">
    function cpFlds() {
      cpVal('vStatus');
      cpVal('vPublTyp');
      cpVal('vPublDatum');
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Hantera tips & trix (Spel)</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
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
      <div class="input"><input type="text" class="fill notnull" name="vTitel" maxlength="100" value="<% = cADD_Titel %>"></div>
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

  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  
  <input type="hidden" name="vStatus_cp">
  <input type="hidden" name="vPublTyp_cp">
  <input type="hidden" name="vPublDatum_cp">
  
  </form>
  
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
        saveDate = Trim(CStr(rsDB("xDatumSparad") & " "))
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