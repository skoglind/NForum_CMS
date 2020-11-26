<% 
  cON_PAGE = "Alla titlar - Titlar - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->

  <%
  If Not GetAcc("CMS202") Then Response.Redirect("/")
  %>
  
  <%
  ' #### FILTER ####
  ' //
  ' ################
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  ' #### ALFALIST ####
    Call GetAlfa(Request.QueryString("alfa"))
    If sAlfa <> Empty Then sAlfaFilter = " WHERE ttText LIKE '" & sAlfa & "%' "
  ' ##################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT *, (SELECT COUNT(*) FROM fsBB_Anv WHERE aTitelID = ttID) AS ttAntalAnv FROM fsBB_Titlar " & sAlfaFilter & " ORDER BY ttSortNr ASC, ttText ASC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = MAXPERPAGE
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Titlar " & sAlfaFilter)(0)
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
    <div class="legend">Alla titlar</div>
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
        <input type="button" value="Ny..." onClick="location.href='_edit.asp?<% = sRebuild %>'">
        <input type="button" value="Radera" onClick="if(confirm('Vill du radera de markerade posterna?')){doSubmit('datalist','a=del');}" <% If Not GetAcc("CMS202") Then Response.Write(" disabled") %>> |
        <select name="a" id="sel_a">
          <option value=""> Fler alternativ... </option>
          <option class="separator" disabled> </option>
        </select>
        <input type="button" value=" » " onClick="doSubmit('datalist','a=' + document.getElementById('sel_a').value);">
      </div>
    </div>
    <div class="alfalist">
      <a href="?<% = sRebuildnoAlfa %>&alfa=" <% If sSendAlfa = Empty Then Response.Write(" class='sel'") %>>Alla</a> |
      <a href="?<% = sRebuildnoAlfa %>&alfa=grind" <% If sSendAlfa = "grind" Then Response.Write(" class='sel'") %>>#</a> |
      <% For zx = 65 To 90 %>
        <a href="?<% = sRebuildnoAlfa %>&alfa=<% = Chr(zx) %>" <% If sSendAlfa = Chr(zx) Then Response.Write(" class='sel'") %>><% = Chr(zx) %></a> |
      <% Next %>
      <a href="?<% = sRebuildnoAlfa %>&alfa=aring" <% If sSendAlfa = "aring" Then Response.Write(" class='sel'") %>>Å</a> |
      <a href="?<% = sRebuildnoAlfa %>&alfa=auml" <% If sSendAlfa = "auml" Then Response.Write(" class='sel'") %>>Ä</a> |
      <a href="?<% = sRebuildnoAlfa %>&alfa=ouml" <% If sSendAlfa = "ouml" Then Response.Write(" class='sel'") %>>Ö</a>
    </div>
    <div class="datalist">
    
    <% If rsDB.EOF Then %>
      <p class="err">Det finns inga poster att visa</p>
    <% End If %>
    
    <form method="post" id="datalist">
      <ul class="list">
        <% For zx = 1 To lMaxPosterPerSida %>
          <% If rsDB.EOF Then Exit For %>
            <% If IsEven Then IsEven = False Else IsEven = True %>
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="background-image: url('/design/icons/papper_1.png')">  <% If rsDB("ttAntalanv") > 0 Then %><span class="number"><strong><% = rsDB("ttAntalanv") %></strong><br>användare</span><% End If %> <input type="checkbox" name="chk_id<% = rsDB("ttID") %>" value="YES"> <a href="_edit.asp?e=<% = rsDB("ttID") %>&<% = sRebuild %>"><% = sEncode(rsDB("ttText")) %></a> <span class="status"><% = sEncode(rsDB("ttForklaring")) %></span> </li>
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
      lStatus_0 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Titlar")(0)
      %>
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> titlar </td></tr>
      </table>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Dina behörigheter</div>
    <div class="inner">
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td> Ser poster</td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Skapa ny </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Redigera </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
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