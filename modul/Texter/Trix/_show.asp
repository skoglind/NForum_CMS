<% 
  cON_PAGE = "Alla tips & trix - Tips & Trix (Spel) - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->

  <%
  If Not GetAcc("CMS111") Then Response.Redirect("/")
  %>
  
  <%
  ' #### FILTER ####
  sFilter = Request.QueryString("f")
  Select Case sFilter
    Case "del"    : sFilter = "del"   : If GetAcc("CMS111") Then sFilterText = "xStatus = 0" Else sFilterText = "NOT xStatus = 0"
    Case "publ"   : sFilter = "publ"  : sFilterText = "xStatus = 4"
    Case "deny"   : sFilter = "deny"  : sFilterText = "xStatus = 3"
    Case "await"  : sFilter = "await" : sFilterText = "xStatus = 2"
    Case Else     : sFilter = ""      : sFilterText = "NOT xStatus = 0"
  End Select
  ' ################
  
  ' #### BEHÖRIGHET ####
  'If NOT GetAcc("CMS111") Then sBFilter = " AND xSkapadAv = " & cCMS_ID
  ' ####################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT * FROM cms_Speltrix LEFT JOIN fsBB_Anv ON fsBB_Anv.aID = cms_Speltrix.xSkapadAv WHERE " & sFilterText & " ORDER BY xStatus ASC, xDatumPublicerad DESC, xDatumSkapad DESC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = MAXPERPAGE
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM cms_Speltrix LEFT JOIN fsBB_Anv ON fsBB_Anv.aID = cms_Speltrix.xSkapadAv WHERE " & sFilterText)(0)
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
    <div class="legend">Alla tips & trix (Spel)</div>
    <div class="editbar">
      <div style="float: right;">
        <form method="get">
          <select name="f">
            <option value=""> Filtrera... </option>
            <option class="separator" disabled> </option>
            <option value="all" class="levelin"> Alla </option>
            <option class="separator" disabled> </option>
            <option value="publ" class="levelin" <% If sFilter = "publ" Then Response.Write(" selected") %>> Publicerade </option>
            <option value="deny" class="levelin" <% If sFilter = "deny" Then Response.Write(" selected") %>> Nekade publicering </option>
            <option value="await" class="levelin" <% If sFilter = "await" Then Response.Write(" selected") %>> Inväntar publicering </option>
          </select>
          <input type="submit" value=" » ">
        </form>
      </div>
      <div style="float: left;">
        <input type="button" value="Ny..." onClick="location.href='_edit.asp?<% = sRebuild %>'">
        <input type="button" value="Radera" onClick="if(confirm('Vill du radera de markerade posterna?')){doSubmit('datalist','a=del');}" <% If sFilter = "del" And Not GetAcc("CMS111") Then Response.Write(" disabled") %>> |
        <select name="a" id="sel_a">
          <option value=""> Fler alternativ... </option>
          <option class="separator" disabled> </option>
        </select>
        <input type="button" value=" » " onClick="doSubmit('datalist','a=' + document.getElementById('sel_a').value);" disabled>
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
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="background-image: url('/design/icons/papper_<% = rsDB("xStatus") %>.png')"> <span class="category"><% = rsDB("xSpelID") %></span> <input type="checkbox" name="chk_id<% = rsDB("xID") %>" value="YES"> <a href="_edit.asp?e=<% = rsDB("xID") %>&<% = sRebuild %>"><% = sEncode(rsDB("xTitel")) %></a> <span class="status">Av: <% = rsDB("aAnvNamn") %></span> </li>
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
    <div class="label">Statistik</div>
    <div class="inner">
      <%
      Set rsStats = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT *, " & _
            "(SELECT COUNT(*) FROM cms_Speltrix WHERE xStatus = 2 " & sBFilter & ") AS iAntalStatus2, " & _
            "(SELECT COUNT(*) FROM cms_Speltrix WHERE xStatus = 3 " & sBFilter & ") AS iAntalStatus3, " & _
            "(SELECT COUNT(*) FROM cms_Speltrix WHERE xStatus = 4 " & sBFilter & ") AS iAntalStatus4 " & _
            "FROM cms_Speltrix"
      rsStats.Open SQL, Con
      
        If rsStats.EOF Then
          lStatus_2 = 0
          lStatus_3 = 0
          lStatus_4 = 0
        Else
          lStatus_2 = rsStats("iAntalStatus2")
          lStatus_3 = rsStats("iAntalStatus3")
          lStatus_4 = rsStats("iAntalStatus4")
        End If
      
      rsStats.Close
      Set rsStats = Nothing
      %>
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td class="td1"> <img src="/design/icons/papper_4_sm.png"> </td><td clasS="td2"> <% = lStatus_4 %> </td><td class="td3"> publicerade </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_2 %> </td><td class="td3"> färdiga </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_3_sm.png"> </td><td clasS="td2"> <% = lStatus_3 %> </td><td class="td3"> nekade</td></tr>
      </table>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Dina behörigheter</div>
    <div class="inner">
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td> Ser poster</td><td> <% If GetAcc("CMS111") Then %>Alla<% Else %>Egna<% End If %> </td></tr>
        <tr><td> Skapa ny </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Redigera </td><td> <% If GetAcc("CMS111") Then %>Alla<% Else %>Egna<% End If %> </td></tr>
        <tr><td> Publicera </td><td> <% If GetAcc("CMS111") Then %>Alla<% ElseIf GetAcc("CMS110") Then %>Egna<% Else %><span style="color: #A00;">Nej</span><% End If %> </td></tr>
        <tr><td> Radera </td><td> <% If GetAcc("CMS111") Then %>Alla<% Else %>Egna<% End If %> </td></tr>
      </table>
    </div>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     