<% 
  cON_PAGE = "Anmälningar - Forum - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->

  <%
  If Not GetAcc("CMS333") Then Response.Redirect("/")
  %>
  
  <%
  ' #### FILTER ####
  ' //
  ' ################
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  ' #### ALFALIST ####

  ' ##################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT *, fsBB_Tradar.tAmne AS anAmne, fsBB_Anv.aAnvNamn AS anAnvadare, fsBB_Tradar.tStatus_Trad AS xEnTrad, fsBB_Tradar.tStatus_Undertrad AS xUnderTrad FROM fsBB_Anmal " & _
        "LEFT JOIN fsBB_Tradar ON tID = anTradID " & _
        "LEFT JOIN fsBB_Anv ON aID = anAnv " & _
        "WHERE anDatum > '" & DATEADD("m", -1, Now) & "' ORDER BY anNoterad ASC, anDatum DESC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = MAXPERPAGE
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anmal WHERE anDatum > '" & DATEADD("m", -1, Now) & "'")(0)
  If Not IsNumeric(lAntalPoster) Or lAntalPoster = Empty Then lAntalPoster = 0
  lAntalSidor = CLng(RoundUp(lAntalPoster, lMaxPosterPerSida))
  
  lPaSida = Request.QueryString("s")
  If Not IsNumeric(lPaSida) Or lPaSida = Empty Then lPaSida = 1
  lPaSida = CLng(lPaSida)
  If lPaSida < 1 Then lPaSida = 1
  If lPaSida > lAntalSidor Then lPaSida = lAntalSidor
  
  If Not lPaSida = 1 And lAntalPoster > 0 Then rsDB.Move (lPaSida - 1) * lMaxPosterPerSida
  ' ################
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa
  sRebuildnoAlfa = "f=" & sFilter & "&s=" & lPaSida
  %>
  
  <div class="datablock rect">
    <div class="legend">Alla anmälningar</div>
    <div class="editbar">
      <div style="float: right;">
        <form method="get">
          <select name="f">
            <option value=""> Filtrera... </option>
            <option class="separator" disabled> </option>
            <option value="all" class="levelin"> Alla </option>
          </select>
          <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
          <input type="submit" value=" » ">
        </form>
      </div>
      <div style="float: left;">
        <input type="button" value="Radera" onClick="if(confirm('Vill du radera de markerade posterna?')){doSubmit('datalist','a=del');}" <% If Not GetAcc("CMS3") Then Response.Write(" disabled") %>> |
        <select name="a" id="sel_a">
          <option value=""> Fler alternativ... </option>
          <option class="separator" disabled> </option>
          <option value="notera"> Markera som behandlad </option>
        </select>
        <input type="button" value=" » " onClick="doSubmit('datalist','a=' + document.getElementById('sel_a').value);">
      </div>
    </div>
    <div class="datalist">
    
    <% If rsDB.EOF Then %>
      <p class="err">Det finns inga anmälningar att visa</p>
    <% End If %>
    
    <form method="post" id="datalist">
      <ul class="list">
        <% For zx = 1 To lMaxPosterPerSida %>
          <% If rsDB.EOF Then Exit For %>
            <% If IsEven Then IsEven = False Else IsEven = True %>
            <%
              tradensID   = rsDB("anTradID")
              bTrad       = rsDB("xEnTrad")
              bUnderTrad  = rsDB("xUndertrad")
              
              If bTrad Then
                TradID  = tradensID
                Go2Trad = tradensID
              Else
                TradID  = bUnderTrad
                Go2Trad = tradensID
              End if
            %>
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="height: 100px; background-image: url('/design/icons/papper_<% If rsDB("anNoterad") Then %>4<% Else %>1<% End If %>.png')"> <input type="checkbox" name="chk_id<% = rsDB("anID") %>" value="YES"> <a href="<% = CMS_SITEADDR %>/avdelning/forum/trad.asp?e=<% = TradID %>&amp;go2=<% = Go2Trad %>" target="_blank"><% = sEncode(rsDB("anAmne")) %></a> <span class="status"><% = sEncode(rsDB("anAnvadare")) %></span> <p><strong>Meddelande:</strong> <% = sEncode(rsDB("anTextM")) %></p> </li>
          <% rsDB.MoveNext %>
        <% Next %>
      </ul>
      <input type="hidden" name="form" value="news">
      <input type="hidden" name="f" value="<% = sFilter %>">
      <input type="hidden" name="s" value="<% = lPaSida %>">
    </form>
    </div>
    <div class="pagebar">
      <% For zx = 1 To lAntalSidor %>
        <a href="?s=<% = zx %>&f=<% = sFilter %>&alfa=<% = sSendAlfa %>" <% If zx = lPaSida Then Response.Write(" style='font-weight: bold;'") %> ><% = zx %></a> <% If Not zx = lAntalSidor Then %>| <% End If %>
      <% Next %>
    </div>
  </div>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    <div class="label">Statistik</div>
    <div class="inner">
      <%
      lStatus_0 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anmal WHERE anNoterad = 0 AND anDatum > '" & DATEADD("m", -1, Now) & "'")(0)
      lStatus_1 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anmal WHERE anNoterad = 1 AND anDatum > '" & DATEADD("m", -1, Now) & "'")(0)
      %>
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> Nya </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_4_sm.png"> </td><td clasS="td2"> <% = lStatus_1 %> </td><td class="td3"> Behandlade </td></tr>
      </table>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Dina behörigheter</div>
    <div class="inner">
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td> Ser poster</td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Notera </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Radera </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
      </table>
    </div>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     