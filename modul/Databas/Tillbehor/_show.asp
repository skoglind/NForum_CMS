<% 
  cON_PAGE = "Alla tillbehör - Tillbehör - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->

  <%
  If Not GetAcc("CMS4") Then Response.Redirect("/")
  %>
  
  <%
  ' #### FILTER ####
  sFilter = Request.QueryString("f")
  If Not IsNumeric(sFilter) Or sFilter = Empty Then sFilter = 0
  
  sQ = Trim(Left(MakeLegal_Large(Request.QueryString("q")), 255))
  
  If CStr(lstKonsol(sFilter)) <> CStr(lstKonsol(0)) Then
    sAddFilter = " AND iKonsol = " & CLng(sFilter)
  Else
    sAddFilter = ""
  End If
  
  If Len(sQ) > 2 Then
    sAddFilter2 = " AND (tTitel LIKE '%" & sQ & "%' OR iNyckelord = '%" & sQ & "%' OR iID IN(SELECT tTillbehorID From cms_TillbehorTitlar WHERE tTitel LIKE '%" & sQ & "%'))"
  Else
    sAddFilter2 = ""
  End If
  
  ' ################
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  ' #### ALFALIST ####
    Call GetAlfa(Request.QueryString("alfa"))
    If sAlfa <> Empty Then sAlfaFilter = "AND tTitel LIKE '" & sAlfa & "%' "
  ' ##################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT *, " & _
        "(SELECT COUNT(*) FROM cms_Bild LEFT JOIN cms_Tillbehortitlar AS iTit ON cms_Bild.bID IN(iTit.tBoxart_BoxFram,iTit.tBoxart_BoxBak,iTit.tBoxart_Tillbehor,iTit.tBoxart_Manual) WHERE tTillbehorID = cms_Tillbehor.iID) AS aAntalBilder " & _
        "FROM cms_Tillbehor LEFT JOIN cms_Tillbehortitlar ON cms_Tillbehor.iStandard_Titel = cms_Tillbehortitlar.tID WHERE 1=1 " & sAlfaFilter & sAddFilter & sAddFilter2 & " ORDER BY tTitel ASC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = MAXPERPAGE
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM cms_Tillbehor LEFT JOIN cms_Tillbehortitlar ON cms_Tillbehor.iStandard_Titel = cms_Tillbehortitlar.tID WHERE 1=1 " & sAlfaFilter & sAddFilter & sAddFilter2)(0)
  If Not IsNumeric(lAntalPoster) Or lAntalPoster = Empty Then lAntalPoster = 0
  lAntalSidor = CLng(RoundUp(lAntalPoster, lMaxPosterPerSida))
  
  lPaSida = Request.QueryString("s")
  If Not IsNumeric(lPaSida) Or lPaSida = Empty Then lPaSida = 1
  lPaSida = CLng(lPaSida)
  If lPaSida < 1 Then lPaSida = 1
  If lPaSida > lAntalSidor Then lPaSida = lAntalSidor
  
  If Not lPaSida = 1 And lAntalPoster > 0 Then rsDB.Move (lPaSida - 1) * lMaxPosterPerSida
  ' ################
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa & "&q=" & sQ
  sRebuildnoAlfa = "f=" & sFilter & "&s=" & lPaSida
  %>
  
  <div class="datablock rect">
    <div class="legend">Alla tillbehör</div>
    <div class="editbar">
      <div style="float: right;">
        <form method="get">
          <select name="f">
            <option value=""> Filtrera... </option>
            <option class="separator" disabled> </option>
            <option value="all" class="levelin"> Alla </option>
            <option class="separator" disabled> </option>
            <% For zx = 1 To lstKonsol(0) %>
              <option value="<% = zx %>" class="levelin" <% If CLng(sFilter) = CLng(zx) Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
            <% Next %>
          </select>
          <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
          <input type="submit" value=" » ">
        </form>
      </div>
      <div style="float: left;">
        <input type="button" value="Ny..." onClick="location.href='_edit.asp?<% = sRebuild %>'">
        <input type="button" value="Radera" onClick="if(confirm('Vill du radera de markerade posterna?')){doSubmit('datalist','a=del');}" <% If Not GetAcc("CMS44") Then Response.Write(" disabled") %>> |
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
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="background-image: url('/design/icons/pad<% If Not rsDB("iSynlig") Then Response.Write("_hid") %>.png')"> <% If rsDB("aAntalBilder") > 0 Then %><span class="number"><strong><% = rsDB("aAntalBilder") %></strong><br>bild(er)</span><% End If %> <input type="checkbox" name="chk_id<% = rsDB("iID") %>" value="YES"> <a href="_edit.asp?e=<% = rsDB("iID") %>&<% = sRebuild %>"><% = sEncode(rsDB("tTitel")) %></a> <span class="status"><% = lstKonsol(rsDB("iKonsol")) %></span> </li>
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
        <a href="?s=<% = zx %>&f=<% = sFilter %>&alfa=<% = sSendAlfa %>&q=<% = sQ %>" <% If zx = lPaSida Then Response.Write(" style='font-weight: bold;'") %> ><% = zx %></a> <% If Not zx = lAntalSidor Then %>| <% End If %>
      <% Next %>
    </div>
  </div>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <div class="databox info">
    <div class="label">Sök</div>
    <div class="inner">
      <form action="_show.asp" method="GET">
        <div class="field">
          <select name="f">
            <%
              lKonsol = sFilter
              If Not IsNumeric(lKonsol) Or lKonsol = Empty Then lKonsol = 0
              lKonsol = CLng(lKonsol)
            %>
            <option value="0"> Alla </option>
            <option class="separator" disabled> </option>
            <% For zx = 1 To lstKonsol(0) %>
              <option value="<% = zx %>" class="levelin" <% If CLng(lKonsol) = zx Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
            <% Next %>
          </select>
        </div>
        <div class="field">
          <input type="text" name="q" value="<% = sEncode(sQ) %>" maxlength="255" style="width: 120px; float: left; color: #333;">
          <input type="submit" value="Sök" style="width: 50px; left;">
        </div>
      </form>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Statistik</div>
    <div class="inner">
      <%
      Dim lCountKonsol(250)
      For zx = 1 To lstKonsolShort(0)
        sListSQL = sListSQL & "(SELECT COUNT(*) FROM cms_Tillbehor WHERE iKonsol = " & CLng(zx) & ") AS iAntalKonsol" & CLng(zx) & ", " 
      Next
      
      Set rsStats = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT *, " & sListSQL & _
            "(SELECT COUNT(*) FROM cms_Tillbehor WHERE kSynlig = 1) AS iAntalStatus0, " & _
            "(SELECT COUNT(*) FROM cms_Tillbehor WHERE kSynlig = 0) AS iAntalStatus1, " & _
            "(SELECT COUNT(*) FROM cms_Tillbehortitlar) AS aAntalTitlar " & _
            "FROM cms_Konsol"
      rsStats.Open SQL, Con
      
        If rsStats.EOF Then
          lStatus_0 = 0
          lStatus_1 = 0
          lAntalTitlar = 0
          For zx = 1 To lstKonsolShort(0)
            lCountKonsol(zx) = 0
          Next
        Else
          lStatus_0 = rsStats("iAntalStatus0")
          lStatus_1 = rsStats("iAntalStatus1")
          lAntalTitlar = rsStats("aAntalTitlar")
          For zx = 1 To lstKonsolShort(0)
            lCountKonsol(zx) = rsStats("iAntalKonsol" & CLng(zx))
          Next
        End If
      
      rsStats.Close
      Set rsStats = Nothing
      %>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <tr><td class="td1"> <img src="/design/icons/pad_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> synliga tillbehör </td></tr>
        <tr><td class="td1"> <img src="/design/icons/pad_hid_sm.png"> </td><td clasS="td2"> <% = lStatus_1 %> </td><td class="td3"> dolda tillbehör </td></tr>
      </table>
      <div class="innerseparator"> </div>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lAntalTitlar %> </td><td class="td3"> tillbehörtitlar </td></tr>
      </table>
      <div class="innerseparator"> </div>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <% For zx = 1 To lstKonsolShort(0) %>
          <tr><td class="td1"> <img src="/design/icons/pad_sm.png"> </td><td class="td2"> <% = lCountKonsol(zx) %> </td><td class="td3"> <a href="?f=<% = zx %>&alfa=<% = sSendAlfa %>"><% = lstKonsolShort(zx) %></a> </td></tr>
        <% Next %>
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
        <tr><td> Radera </td><td> <% If GetAcc("CMS44") Then %><span style="color: #0A0;">Ja</span><% Else %><span style="color: #A00;">Nej</span><% End If %> </td></tr>
      </table>
    </div>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     