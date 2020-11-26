<% 
  cON_PAGE = "Hantera spelgrupp - Spelgrupper - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS4") Then Response.Redirect("/")
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
    SQL = "SELECT * FROM cms_Spelserier WHERE ssID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("ssID")
      cADD_Namn               = sEncode(rsDB("ssNamn"))
      cADD_Text               = sEncode(rsDB("ssTextM"))
      cADD_Serien             = rsDB("ssSerien")
      cADD_Nyckelord          = sEncode(rsDB("ssNyckelord"))
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
    <div class="legend">Hantera spelgrupp</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Namn</div>
      <div class="input"><input type="text" class="fill notnull" name="vNamn" maxlength="255" value="<% = cADD_Namn %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Spelserien</div>
      <div class="input"><input type="checkbox" class="fill" name="vSerien" maxlength="255" value="YES" style="width: 25px;" <% If cADD_Serien Then Response.Write(" checked") %>></div>
    </div>
    <div class="in_text"> <p>Kryssa i denna ruta om du åtsyftar hela spelserien (tex Mega Man), är det bara en underserie (tex Mega Man X) kryssa inte i den.</p> </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Nyckelord</div>  
      <div class="input"><input type="text" class="fill" name="vNyckelord" maxlength="500" value="<% = cADD_Nyckelord %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter grupper. Seprarera varje nyckelord med ett mellanslag.</p> </div>
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
        saveDate = Trim(CStr(rsDB("ssDatumSparad") & " "))
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