<% 
  cON_PAGE = "Från lista - Snabblistning av spel - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS4") Then Response.Redirect("/")
  %>

  <%
  
  %>
  
  <script type="text/javascript">
    function cpFlds() {
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
            <option value="<% = zx %>"> <% = lstKonsol(zx) %> </option>
          <% Next %>
        </select>
      </div>
    </div>
  </div>
  
  <div class="datablock rect morepadding">
    <div class="legend">Från lista - Snabblistning av spel</div>

    <div class="in_row">
      <div class="texttools">
        Lista med spel, gör en radbrytning mellan varje
      </div>
      <textarea id="myText" name="vText"><% = cADD_Text %></textarea>
    </div>
    
  </div>
  
  </form>

  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    
    <div class="inner" style="text-align: center;">
      <input onclick="cpFlds();saveform('em',0);" name="savebtn" class="save" type="button" value="Läs in..." <% If bCantSave Then Response.Write(" disabled") %>>
      <input style="display: none;" onclick="cpFlds();saveform('em',1);" name="savebtn" class="save_continue" type="button" value="Spara och fortsätt..." <% If bCantSave Then Response.Write(" disabled") %>>
      <input style="display: none;" onclick="cpFlds();saveform('em',2);" name="savebtn" class="save_return" type="button" value="Spara och återgå..." <% If bCantSave Then Response.Write(" disabled") %>>
      <input style="display: none;" onclick="location.href='_show.asp?<% = sRebuild %>';" name="savebtn" class="cancel" type="button" value="Avbryt">
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
      
      <iframe name="processbox" id="processbox" style="width: 174px; height: 180px; display: block;" frameborder=0 src="/_awaiting.asp"></iframe>
    </div>
  </div>
  
  <div class="databox info" style="display: none;">
    <div class="label">Sparad senast</div>
    <div class="inner">
      <div class="radio" style="background-image: url('/design/icons/radio_<% If isSaved Then Response.Write("true") Else Response.Write("false") %>.png');" id="savedstatus"><% = saveDate %></div>
    </div>
  </div>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     