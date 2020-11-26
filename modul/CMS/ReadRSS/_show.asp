<% 
  cON_PAGE = "Alla RSS-Strömmar - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  ' #### FILTER ####
  ' ################
  
  ' #### BEHÖRIGHET ####
  ' ####################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT * FROM rss_Data LEFT JOIN rss_Feed ON rfID = rdFeedID ORDER BY rdDate DESC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = 50
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM rss_Data")(0)
  If Not IsNumeric(lAntalPoster) Or lAntalPoster = Empty Then lAntalPoster = 0
  lAntalSidor = CLng(RoundUp(lAntalPoster, lMaxPosterPerSida))
  
  lPaSida = Request.QueryString("s")
  If Not IsNumeric(lPaSida) Or lPaSida = Empty Then lPaSida = 1
  lPaSida = CLng(lPaSida)
  If lPaSida < 1 Then lPaSida = 1
  If lPaSida > lAntalSidor Then lPaSida = lAntalSidor
  
  If Not lPaSida = 1 And lAntalPoster > 0 Then rsDB.Move (lPaSida - 1) * lMaxPosterPerSida
  ' ################
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida
  %>
  
  <div class="datablock rect">
    <div class="legend">Alla RSS-Strömmar</div>
    <div class="editbar">
      <div style="float: right;">
        <form method="get">
          <select name="f">
            <option value=""> Filtrera... </option>
          </select>
          <input type="submit" value=" » " disabled>
        </form>
      </div>
      <div style="float: left;">
        <input type="button" value="Ny..." disabled>
        <input type="button" value="Radera" disabled> |
        <select name="a" id="sel_a">
          <option value=""> Fler alternativ... </option>
        </select>
        <input type="button" value=" » " disabled>
      </div>
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
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="background-image: url('/design/icons/papper_4.png')"> <span class="category"><% = rsDB("rdDate") %></span> <input type="checkbox" name="chk_id<% = rsDB("rdID") %>" value="YES" disabled> <a href="<% = rsDB("rdURL") %>" target="_blank" title="<% = rsDB("rdTitle") %>"><% = CutWord(rsDB("rdTitle"),60) %></a> <span class="status"><strong><% = rsDB("rfName") %></strong></span> </li>
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
        <a href="?s=<% = zx %>&f=<% = sFilter %>" <% If zx = lPaSida Then Response.Write(" style='font-weight: bold;'") %> ><% = zx %></a> <% If Not zx = lAntalSidor Then %>| <% End If %>
      <% Next %>
    </div>
  </div>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    <div class="label">Dina behörigheter</div>
    <div class="inner">
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td> Ser poster</td><td> Alla </td></tr>
      </table>
    </div>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     