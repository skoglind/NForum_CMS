<% 
  cON_PAGE = "Hantera företag - Företag - CMS"
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
    SQL = "SELECT * FROM cms_Foretag WHERE fID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Hemsida            = "http://"
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("fID")
      cADD_Namn               = sEncode(rsDB("fNamn"))
      cADD_VD                 = sEncode(rsDB("fVD"))
      cADD_Text               = sEncode(rsDB("fTextM"))
      cADD_Nyckelord          = sEncode(rsDB("fNyckelord"))
      cADD_Hemland            = sEncode(rsDB("fHemland"))
      cADD_Blevtill           = sEncode(rsDB("fBlevtill"))
      cADD_Hemsida            = sEncode(rsDB("fHemsida"))
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
    <div class="legend">Hantera företag</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Namn</div>
      <div class="input"><input type="text" class="fill notnull" name="vNamn" maxlength="255" value="<% = cADD_Namn %>"></div>
    </div>
    <div class="in_text"> <p>Ange hela företagsnamnet, men förkorta följande ord: <strong>Corporation/Company = Co.</strong>, <strong>Limited = Ltd.</strong>, <strong>Incorporated = Inc.</strong>.</p> </div>
    
    <div class="in_row">
      <div class="text">VD / CEO</div>
      <div class="input"><input type="text" class="fill" name="vVD" maxlength="255" value="<% = cADD_VD %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Hemsida</div>
      <div class="input"><input type="text" class="fill" name="vHemsida" maxlength="255" value="<% = cADD_Hemsida %>"></div>
    </div>
    <div class="in_text"> <p>Fyll i företagets hemsida. Finns en sida på svenska anger du den annars fyller du i den engelska. Finns den inte på något av de angivna språken fyll då i deras huvudsida.</p> </div>
    
    <div class="in_row">
      <div class="text">Stad, Land</div>
      <div class="input"><input type="text" class="fill" name="vHemland" maxlength="255" value="<% = cADD_Hemland %>"></div>
    </div>
    
    <div class="in_text"> <p>Fyll in i vilket land som företaget är lokaliserat. Om företaget finns i flera land skriv in det land där de startade. Fyll i på följande sätt <strong>"Kyoto, Japan"</strong>. Ligger det i USA så ange även stat enligt följande <strong>"Redmond, Washington, USA"</strong>.</p> </div>
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Nyckelord</div>  
      <div class="input"><input type="text" class="fill" name="vNyckelord" maxlength="500" value="<% = cADD_Nyckelord %>"></div>
    </div>
    
    <div class="in_text"> <p>Skriv gärna in några nyckelord för att underlätta för sökning efter företag. Seprarera varje nyckelord med ett mellanslag.</p> </div>
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
  
  <div class="databox info">
    <div class="label">Logotype</div>
    <div class="inner" id="statusarea" style="text-align: center;">
      --
    </div>
  </div>
  
  <% If Not rsDB.Eof And GetAcc("CMS44") Then %>
    <div class="databox info">
      <div class="label">Dublettrensning</div>
      <div class="inner" id="statusarea" style="text-align: center;" >
        <form method="POST" id="glueform" name="glueform">
          <p>Slå ihop med följande:</p>
          <div class="field">
            <input type="text" class="fill" name="fAdd_Name" id="fAdd_Name" value="Inget företag valt." readonly style="width: 110px;">
            <input type="button" value="Välj..." onclick="showPicker('/picker/cms_Foretag.asp','fAdd_ID','fAdd_Name');" style="width: 56px;">
            <input type="hidden" name="fAdd_ID" id="fAdd_ID">
            <input type="hidden" name="fMy_ID" id="fMy_ID" value="<% = cADD_ID %>">
          </div>
          <div class="innerseparator"> </div>
          <div class="field"><input type="button" value="Slå ihop och ta bort denna" style="font-weight: bold;" onclick="if(confirm('Vill du verkligen TA BORT DETTA FÖRETAG och slå ihop det\n med ovan angivna företag? Åtgärden kan INTE ångras.')){gluesave();}"></div>
        </form>
      </div>
    </div>
  <% End If %>
  
  <script type="text/javascript">getPage("__innerfld.asp?e=<% = cADD_ID %>");</script>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     