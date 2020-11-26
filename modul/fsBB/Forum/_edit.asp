<% 
  cON_PAGE = "Hantera forumkategori - Forum - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS333") Then Response.Redirect("/")
  %>

  <%
  lID = Request.QueryString("e")
  If Not IsNumeric(lID) Or lID = Empty Then lID = 0
  lID = CLng(lID)
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM fsBB_Forum WHERE fID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("fID")
      cADD_Namn               = sEncode(rsDB("fName"))
      cADD_Info               = sEncode(rsDB("fInfo"))
      cADD_Color              = sEncode(rsDB("fColor"))
      cADD_Icon               = sEncode(rsDB("fIcon"))
      cADD_SortNr             = CLng(rsDB("fSortNr"))
      cADD_Sortering          = CLng(rsDB("fSortering"))
      
      cADD_bNoAllView         = rsDB("fNoAllView")
      cADD_bSplitter          = rsDB("fSplitterBefore")
      cADD_bGroup             = rsDB("fGroup")
      cADD_bHideMe            = rsDB("fHideMe")
      
      cADD_lNewThread         = rsDB("fSec_NewThread")
      cADD_lNewReply          = rsDB("fSec_NewReply")
      cADD_lView              = rsDB("fSec_View")
      cADD_lModerator         = rsDB("fSec_Mod")
      cADD_lGroupForums       = rsDB("fGroupForums")
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  Call GetAlfa(Request.QueryString("alfa"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa
  ' ##################
  %>
  
  <script type="text/javascript">
    function cpFlds() {
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Hantera forumkategori</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Namn</div>
      <div class="input"><input type="text" class="fill notnull" name="vNamn" maxlength="30" value="<% = cADD_Namn %>"></div>
    </div>
    
    <div class="in_row">
      <div class="text">Information</div>
      <div class="input"><input type="text" class="fill" name="vInfo" maxlength="255" value="<% = cADD_Info %>"></div>
    </div>
    <div class="in_text"> <p>Detaljerad beskrivning av forumet.</p> </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Sortering</div>
      <div class="input">
        <select name="vSortering">
          <option value="0" <% If CLng(cADD_Sortering) = 0 Then Response.Write(" selected") %>> Senast uppdaterad, överst </option>
          <option value="1" <% If CLng(cADD_Sortering) = 1 Then Response.Write(" selected") %>> Senast skapad, överst </option>
          <option value="2" <% If CLng(cADD_Sortering) = 2 Then Response.Write(" selected") %>> Efter ämnestitel </option>
        </select>
      </div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Placering</div>
      <div class="input"><input type="text" class="fill" name="vSortNr" maxlength="3" value="<% = cADD_SortNr %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vNoAllView" value="YES" <% If cADD_bNoAllView Then Response.Write(" checked") %>> Visa ej i alla</div>
      <div class="input">Döljer forumets trådar från "Alla forum".</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vSplitterBefore" value="YES" <% If cADD_bSplitter Then Response.Write(" checked") %>> Splitter innan</div>
      <div class="input">Placerar en spärrlinje innan forumet i forumindexet.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vHideMe" value="YES" <% If cADD_bHideMe Then Response.Write(" checked") %>> Dölj mig</div>
      <div class="input">Döljer forumet.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vGroup" value="YES" <% If cADD_bGroup Then Response.Write(" checked") %>> Gruppforum</div>
      <div class="input">Forumet grupperar andra forum.</div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="twopart">Skapa trådar:</div>
      <div class="twopart">Skapa inlägg:</div>
    </div>
    
    <div class="in_row">
      <% Set rsList = Server.CreateObject("ADODB.Recordset") %>
    
      <% 
        SQL = "SELECT * FROM fsBB_Titlar ORDER BY ttSortNr ASC, ttText ASC"
        rsList.Open SQL, con
      %>
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="newthread_radio" value="0" id="newthread" onclick="getToggle('newthread');"></div> <div class="lbl" onclick="toggletrue('newthread');getToggle('newthread');">Alla</div> </div>
          <div class="row"> <div class="chk"><input type="checkbox" name="newthread" value="99" id="newthread99" onclick="setToggle('newthread');" <% If InStr(cADD_lNewThread,";99;") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('newthread99');setToggle('newthread');">Ingen</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="newthread" value="<% = rsList("ttID") %>" id="newthread<% = rsList("ttID") %>" onclick="setToggle('newthread');" <% If InStr(cADD_lNewThread,";" & rsList("ttID") & ";") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('newthread<% = rsList("ttID") %>');setToggle('newthread');"><% = rsList("ttForklaring") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
        SQL = "SELECT * FROM fsBB_Titlar ORDER BY ttSortNr ASC, ttText ASC"
        rsList.Open SQL, con
      %>
      
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="newreply_radio" value="0" id="newreply" onclick="getToggle('newreply');"></div> <div class="lbl" onclick="toggletrue('newreply');getToggle('newreply');">Alla</div> </div>
          <div class="row"> <div class="chk"><input type="checkbox" name="newreply" value="99" id="newreply99" onclick="setToggle('newreply');" <% If InStr(cADD_lNewReply,";99;") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('newreply99');setToggle('newreply');">Ingen</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="newreply" value="<% = rsList("ttID") %>" id="newreply<% = rsList("ttID") %>" onclick="setToggle('newreply');" <% If InStr(cADD_lNewReply,";" & rsList("ttID") & ";") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('newreply<% = rsList("ttID") %>');setToggle('newreply');"> <% = rsList("ttForklaring") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
      %>
      
      <% Set rsList = Nothing %>
    </div>

    <div class="in_row">
      <div class="twopart">Visa forum:</div>
      <div class="twopart">Moderator:</div>
    </div>
    
    <div class="in_row">
      <% Set rsList = Server.CreateObject("ADODB.Recordset") %>
    
      <% 
        SQL = "SELECT * FROM fsBB_Titlar ORDER BY ttSortNr ASC, ttText ASC"
        rsList.Open SQL, con
      %>
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="view_radio" value="0" id="view" onclick="getToggle('view');"></div> <div class="lbl" onclick="toggletrue('view');getToggle('view');">Alla</div> </div>
          <div class="row"> <div class="chk"><input type="checkbox" name="view" value="99" id="view99" onclick="setToggle('view');" <% If InStr(cADD_lView,";99;") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('view99');setToggle('view');">Ingen</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="view" value="<% = rsList("ttID") %>" id="view<% = rsList("ttID") %>" onclick="setToggle('view');" <% If InStr(cADD_lView,";" & rsList("ttID") & ";") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('view<% = rsList("ttID") %>');setToggle('view');"><% = rsList("ttForklaring") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
        SQL = "SELECT * FROM fsBB_Titlar ORDER BY ttSortNr ASC, ttText ASC"
        rsList.Open SQL, con
      %>
      
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="mod_radio" value="0" id="mod" onclick="getToggle('mod');"></div> <div class="lbl" onclick="toggletrue('mod');getToggle('mod');">Alla</div> </div>
          <div class="row"> <div class="chk"><input type="checkbox" name="mod" value="99" id="mod99" onclick="setToggle('mod');" <% If InStr(cADD_lModerator,";99;") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('mod99');setToggle('mod');">Ingen</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="mod" value="<% = rsList("ttID") %>" id="mod<% = rsList("ttID") %>" onclick="setToggle('mod');" <% If InStr(cADD_lModerator,";" & rsList("ttID") & ";") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('mod<% = rsList("ttID") %>');setToggle('mod');"> <% = rsList("ttForklaring") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <%
        rsList.Close
      %>
      
      <% Set rsList = Nothing %>
      
      <script type="text/javascript">
        setToggle('newthread');
        setToggle('newreply');
        setToggle('view');
        setToggle('mod');
      </script>
    </div>
    
    <div class="in_row">
      <div class="twopart">Grupperat forum:</div>
      <div class="twopart">&nbsp;</div>
    </div>
    
    <div class="in_row">
      <% Set rsList = Server.CreateObject("ADODB.Recordset") %>
    
      <% 
        SQL = "SELECT * FROM fsBB_Forum WHERE fGroup = 0 AND fID <> " & CLng(cADD_ID) & " ORDER BY fSortNr ASC, fName ASC"
        rsList.Open SQL, con
      %>
      <div class="chklist_outer">
        <div class="chklist">
          <div class="row" style="background-color: #EEE;"> <div class="chk"><input type="radio" name="group_radio" value="0" id="group" onclick="getToggle('group');"></div> <div class="lbl" onclick="toggletrue('group');getToggle('group');">Inga forum</div> </div>
          <% Do Until rsList.EOF %>
            <div class="row"> <div class="chk"><input type="checkbox" name="group" value="<% = rsList("fID") %>" id="group<% = rsList("fID") %>" onclick="setToggle('group');" <% If InStr("," & cADD_lGroupForums & ",", "," & rsList("fID") & ",") Then Response.Write(" checked") %>></div> <div class="lbl" onclick="toggle('group<% = rsList("fID") %>');setToggle('group');"><% = rsList("fName") %></div> </div>
            <% rsList.MoveNext %>
          <% Loop %>
        </div>
      </div>
      
      <script type="text/javascript">
        setToggle('group');
      </script>
    </div>
  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
  
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
        saveDate = Trim(CStr(rsDB("fDatumSparad") & " "))
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
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     