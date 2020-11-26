<% 
  cON_PAGE = "Hantera omröstning - Omröstningar - CMS"
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
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * " & _
          "FROM cms_Omrostning_Fraga " & _ 
          "WHERE omfID = " & CLng(lID)
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID                 = rsDB("omfID")
      cADD_Fraga              = sEncode(rsDB("omfFraga"))
      cADD_SlutDatum          = FormatDateTime((rsDB("omfSlutDatum")), vbShortDate)
      cADD_Synlig             = rsDB("omfSynlig")
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
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Hantera omröstning</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Fråga</div>
      <div class="input"><input type="text" class="fill notnull" name="vFraga" maxlength="255" value="<% = cADD_Fraga %>"></div>
    </div>
    
    <div class="in_row">
      <div class="text">Slutdatum</div>
      <div class="input"><input type="text" class="fill notnull" name="vSlutdatum" maxlength="10" value="<% = cADD_SlutDatum %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Synlig</div>
      <div class="input"><input type="checkbox" class="fill" name="vSynlig" maxlength="255" value="YES" style="width: 25px;" <% If cADD_Synlig Then Response.Write(" checked") %>></div>
    </div>
  
  </div>
  
  <!-- DYNAROWS -->
    <!-- DIV ATT KLONA -->
    <div id="rowclone" style="display: none;">
      <div style="float: left; width: 570px; margin: 0 1px 0 1px; border-bottom: dotted 1px #CCC; padding: 4px 2px 2px 2px;">
        <div style="float: left;">Val: <input id="vVal_XXXX" name="vVal_XXXX" type="text" maxlength=50 style="width: 350px;"></div>
        <div style="float: left; margin-left: 5px;">Sortnr: <input id="vSortNr_XXXX" name="vSortNr_XXXX" value=0 type="text" maxlength=2 style="width: 25px; text-align: center;"></div>
        <div style="float: right;">
          <input type="button" value="Ta bort" onclick="if(confirm('Vill du ta bort raden?')){delRow(XXXX);}">
        </div>
      </div>
    </div>
    <!-- /RAD ATT KLONA -->
    
    <div class="datablock rect morepadding">
      <div class="legend">Omröstningsalternativ<input type="button" value="Nytt alternativ" onclick="addRow(getSlumpID());" style="float: right;"></div>
      <div id="allrows" class="lista">
        <!-- HÄR HAMNAR KLONEN -->
      </div>
    </div>
    
    <!-- HÄMTA RADER FRÅN DATABASEN -->
    <script type="text/javascript">
    <%
      Set rsVal = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Omrostning_Val WHERE omvFraga = " & CLng(cADD_ID)
      rsVal.Open SQL, Con, 1, 3
      
        vID = 0
        Do Until rsVal.EOF
          vID = rsVal("omvID")
          %>  
          addRow(<% = vID %>);
          setRowData(<% = vID %>,"vVal","<% = rsVal("omvText") %>");
          setRowData(<% = vID %>,"vSortNr","<% = rsVal("omvSortNr") %>");
          <%
          rsVal.MoveNext
        Loop
      
      rsVal.Close
    %>
    </script>
    <!-- /HÄMTA RADER FRÅN DATABASEN -->
  <!-- /DYNAROWS -->
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  
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
        saveDate = Trim(CStr(rsDB("omfDatumSparad") & " "))
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
  
  <script type="text/javascript">getPage("__innerfld.asp?e=<% = cADD_ID %>");</script>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     