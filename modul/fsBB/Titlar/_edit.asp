<% 
  cON_PAGE = "Hantera titel - Titlar - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS202") Then Response.Redirect("/")
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
    SQL = "SELECT * FROM fsBB_Titlar WHERE ttID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("ttID")
      cADD_Text               = sEncode(rsDB("ttText"))
      cADD_Forklaring         = sEncode(rsDB("ttForklaring"))
      cADD_SortNr             = CLng(rsDB("ttSortNr"))
      cADD_bSystem            = rsDB("ttSystem")
      cADD_bAdmin             = rsDB("ttAdmin")
      cADD_bRed               = rsDB("ttRed")
      cADD_bModerator         = rsDB("ttModerator")
      cADD_bVIP               = rsDB("ttVIP")
      cADD_bPlus              = rsDB("ttPlus")
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
    <div class="legend">Hantera titel</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Text</div>
      <div class="input"><input type="text" class="fill notnull" name="vText" maxlength="30" value="<% = cADD_Text %>"></div>
    </div>
    <div class="in_text"> <p>Detta är den synliga delen i till exempel forumet och på profilen.</p> </div>
    
    <div class="in_row">
      <div class="text">Förklaring</div>
      <div class="input"><input type="text" class="fill" name="vForklaring" maxlength="150" value="<% = cADD_Forklaring %>"></div>
    </div>
    <div class="in_text"> <p>Detaljerad beskrivning av titeln.</p> </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Sortering</div>
      <div class="input"><input type="text" class="fill" name="vSortNr" maxlength="3" value="<% = cADD_SortNr %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vSystem" value="YES" <% If cADD_bSystem Then Response.Write(" checked") %>> System</div>
      <div class="input">Systemanvändare, Tha MASTER.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vAdmin" value="YES" <% If cADD_bAdmin Then Response.Write(" checked") %>> Admin</div>
      <div class="input">Adminanvändare, användaren blir administratör.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vRed" value="YES" <% If cADD_bRed Then Response.Write(" checked") %>> Red</div>
      <div class="input">Redaktör, denna fyller ingen funktion.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vModerator" value="YES" <% If cADD_bModerator Then Response.Write(" checked") %>> Moderator</div>
      <div class="input">Moderator, denna fyller ingen funktion.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vVIP" value="YES" <% If cADD_bVIP Then Response.Write(" checked") %>> VIP</div>
      <div class="input">VIP-Användare, låter användaren komma åt VIP-avdelningar.</div>
    </div>
    
    <div class="in_row">
      <div class="text"><input type="checkbox" name="vPlus" value="YES" <% If cADD_bPlus Then Response.Write(" checked") %>> Plus</div>
      <div class="input">Plusanvändare, ger en användare lite extra behörigheter.</div>
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
        saveDate = Trim(CStr(rsDB("ttDatumSparad") & " "))
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