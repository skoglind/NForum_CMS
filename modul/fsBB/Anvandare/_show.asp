<% 
  cON_PAGE = "Alla användare - Användare - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->

  <%
  If Not GetAcc("CMS2") Then Response.Redirect("/")
  %>
  
  <%
  ' #### FILTER ####
  sFilter = Request.QueryString("f")
  Select Case sFilter
    Case "reg_1"    : sFilter = "reg_1"  : sFilterText = "aAktiverad = 1 AND aMedlemSedan > '" & DateAdd("d", -14, Now) & "'"
    Case "reg_2"    : sFilter = "reg_2"  : sFilterText = "aAktiverad = 1 AND aMedlemSedan > '" & DateAdd("m", -1, Now) & "'"
    Case "unact"    : sFilter = "unact"  : sFilterText = "aAktiverad = 0"
    Case "ban"      : sFilter = "ban"    : sFilterText = "aAktiverad = 1 AND aBlockadTill >= '" & Now & "'"
    Case "cms"      : sFilter = "cms"    : sFilterText = "aAktiverad = 1 AND aBlockadTill < '" & Now & "' AND aS_CMS = 1"
    Case Else       : sFilter = ""       : sFilterText = "aAktiverad = 1 AND aBlockadTill < '" & Now & "'"
  End Select
  ' ################
  
  ' #### BEHÖRIGHET ####
  ' //
  ' ####################
  
  ' #### ALFALIST ####
    Call GetAlfa(Request.QueryString("alfa"))
    If sAlfa <> Empty Then sAlfaFilter = " AND aAnvNamn LIKE '" & sAlfa & "%' "
  ' ##################
  
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT * FROM fsBB_Anv LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID WHERE " & sFilterText & sAlfaFilter & " ORDER BY aAnvNamn ASC"
  rsDB.Open SQL, Con
  
  ' #### PAGING ####
  lMaxPosterPerSida = MAXPERPAGE
  lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE " & sFilterText & sAlfaFilter)(0)
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
    <div class="legend">Alla användare</div>
    <div class="editbar">
      <div style="float: right;">
        <form method="get">
          <select name="f">
            <option value=""> Filtrera... </option>
            <option class="separator" disabled> </option>
            <option value="all" class="levelin"> Alla </option>
            <option value="reg_1" class="levelin" <% If sFilter = "reg_1" Then Response.Write(" selected") %>> Registrerade inom (2 Veckor) </option>
            <option value="reg_2" class="levelin" <% If sFilter = "reg_2" Then Response.Write(" selected") %>> Registrerade inom (1 Månad) </option>
            <option value="unact" class="levelin" <% If sFilter = "unact" Then Response.Write(" selected") %>> Ej aktiverade </option>
            <option value="ban" class="levelin" <% If sFilter = "ban" Then Response.Write(" selected") %>> Bannade </option>
            <option class="separator" disabled> </option>
            <option value="cms" class="levelin" <% If sFilter = "cms" Then Response.Write(" selected") %>> CMS Aktiverade </option>
          </select>
          <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
          <input type="submit" value=" » ">
        </form>
      </div>
      <div style="float: left;">
        <input type="button" value="Ny..." onClick="location.href='_edit.asp?<% = sRebuild %>'" <% If Not GetAcc("CMS202") Then Response.Write(" disabled") %>>
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
            <%
            If rsDB("aAktiverad") = False Then
              sUserIkon = 1
            ElseIf rsDB("aBlockadTill") >= Now Then
              sUserIkon = 2
            ElseIf rsDB("aS_CMS") = True Then
              sUserIkon = 3
            Else
              sUserIkon = 0
            ENd If
            %>
            <li <% If IsEven Then Response.Write(" class='IO'") %> style="background-image: url('/design/icons/user_<% = sUserIkon %>.png')"> <span class="category"><% = rsDB("ttText") %></span> <input type="checkbox" name="chk_id<% = rsDB("aID") %>" value="YES"> <a href="_edit.asp?e=<% = rsDB("aID") %>&<% = sRebuild %>"><% = sEncode(rsDB("aAnvNamn")) %></a> <span class="status"><% = sEncode(rsDB("aNamn")) %></span> </li>
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
      Set rsStats = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT *, " & _
            "(SELECT COUNT(*) FROM fsBB_Anv WHERE aAktiverad = 1 AND aBlockadTill < '" & Now & "') AS iAntalStatus0, " & _
            "(SELECT COUNT(*) FROM fsBB_Anv WHERE aAktiverad = 0) AS iAntalStatus1, " & _
            "(SELECT COUNT(*) FROM fsBB_Anv WHERE aAktiverad = 1 AND aBlockadTill >= '" & Now & "') AS iAntalStatus2, " & _
            "(SELECT COUNT(*) FROM fsBB_Anv WHERE aAktiverad = 1 AND aBlockadTill < '" & Now & "' AND aS_CMS = 1) AS iAntalStatus3 " & _
            "FROM fsBB_Anv"
      rsStats.Open SQL, Con
      
        If rsStats.EOF Then
          lStatus_0 = 0
          lStatus_1 = 0
          lStatus_2 = 0
          lStatus_3 = 0
        Else
          lStatus_0 = rsStats("iAntalStatus0")
          lStatus_1 = rsStats("iAntalStatus1")
          lStatus_2 = rsStats("iAntalStatus2")
          lStatus_3 = rsStats("iAntalStatus3")
        End If
      
      rsStats.Close
      Set rsStats = Nothing
      %>
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td class="td1"> <img src="/design/icons/user_0_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> användare </td></tr>
        <tr><td class="td1"> <img src="/design/icons/user_1_sm.png"> </td><td clasS="td2"> <% = lStatus_1 %> </td><td class="td3"> inväntar aktivering </td></tr>
        <tr><td class="td1"> <img src="/design/icons/user_2_sm.png"> </td><td clasS="td2"> <% = lStatus_2 %> </td><td class="td3"> bannade</td></tr>
        <tr><td class="td1"> <img src="/design/icons/user_3_sm.png"> </td><td clasS="td2"> <% = lStatus_3 %> </td><td class="td3"> CMS användare</td></tr>
      </table>
    </div>
  </div>
  
  <div class="databox info">
    <div class="label">Dina behörigheter</div>
    <div class="inner">
      <table class="list" cellpadding=0 cellspacing=0>
        <tr><td> Ser användare</td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Skapa ny </td><td> <% If GetAcc("CMS202") Then %><span style="color: #0A0;">Ja</span><% Else %><span style="color: #A00;">Nej</span><% End If %> </td></tr>
        <tr><td> Redigera </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Sätta behörigheter </td><td> <% If GetAcc("CMS202") Then %><span style="color: #0A0;">Ja</span><% Else %><span style="color: #A00;">Nej</span><% End If %> </td></tr>
        <tr><td> Banna </td><td> <span style="color: #0A0;">Ja</span> </td></tr>
        <tr><td> Banna administratörer </td><td> <% If GetAcc("CMS202") Then %><span style="color: #0A0;">Ja</span><% Else %><span style="color: #A00;">Nej</span><% End If %> </td></tr>
      </table>
    </div>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
      
<!--#INCLUDE FILE="../../../_defbottom.asp"-->     