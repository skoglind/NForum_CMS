<% 
  cON_PAGE = "Hantera användare - Användare - CMS"
%>

<!--#INCLUDE FILE="../../../_deftop.asp"-->
  
  <%
  If Not GetAcc("CMS2") Then Response.Redirect("/")
  %>

  <%
  lID = Request.QueryString("e")
  If Not IsNumeric(lID) Or lID = Empty Then lID = 0
  lID = CLng(lID)
  
  ' #### BEHÖRIGHET ####
  If NOT GetAcc("CMS202") Then sBFilter = ""
  ' ####################
  
  Con_Open
  
  ' #### LADDA IN DATA ####
    Set rsDB = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM fsBB_Anv LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID WHERE aID = " & CLng(lID) & sBFilter
    rsDB.Open SQL, Con
    
    If rsDB.EOF Then        ' NY POST
      If Not GetAcc("CMS202") Then Response.Redirect("_show.asp")
    
      lPBStatus = "NewPost"
      
      cADD_ID                 = 0
      cADD_Titel              = FSBB_DEFAULTUSER
    Else                    ' EDITERAD POST
      lPBStatus = "EditPost"
      
      cADD_ID             = rsDB("aID")
      cADD_Titel          = rsDB("aTitelID")
      cADD_EgenTitel      = sEncode(rsDB("aEgenTitel"))
      cADD_AnvNamn        = sEncode(rsDB("aAnvNamn"))
      cADD_Epost          = sEncode(rsDB("aEpost"))
      cADD_Namn           = sEncode(rsDB("aNamn"))
      cADD_Plats          = sEncode(rsDB("aPlats"))
      cADD_Hemsida        = sEncode(rsDB("aHemsida"))
      cADD_MSN            = sEncode(rsDB("aMSN"))
      cADD_ICQ            = sEncode(rsDB("aICQ"))
      cADD_Signatur       = sEncode(rsDB("aSignatur"))
      cADD_Profil         = sEncode(rsDB("aPM"))
      cADD_Avatar         = FSBB_AVATARER & "u" & Right("000000" & cADD_ID, 6) & ".jpg"
      cADD_Avatar_URL     = FSBB_AVATARER_URL & "u" & Right("000000" & cADD_ID, 6) & ".jpg"
      cADD_HasCMS         = rsDB("aS_CMS")
      cADD_Ratter         = rsDB("aS_CMSRatter")
      cADD_BlockStatus    = rsDB("aBlockStatus")
      cADD_IsActiv        = rsDB("aAktiverad")
      cADD_TimeStamp      = rsDB("aTimeStamp")
      cADD_LoggaUt        = rsDB("aLOCK")
      
      If cADD_TimeStamp > DateAdd("n", -5, Now) Then cADD_Online = True
    End If
  ' ##################
  
  ' #### REMEMBER ####
  sFilter = noFnutt(Request.QueryString("f"))
  lPaSida = noFnutt(Request.QueryString("s"))
  Call GetAlfa(Request.QueryString("alfa"))
  
  sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa
  ' ##################
  
  If rsDB.EOF Then If GetAcc("CMS202") Then bCanCreateNew = True
  %>
  
  <script type="text/javascript">
    function cpFlds() {
      cpVal('vBannad');
      <% If GetAcc("CMS202") Then %>
        cpVal('vCMS');
        cpVal('vCMS0');
        cpVal('vCMS1');
        <% If CLng(cADD_ID) <> CLng(cCMS_ID) Then %>cpVal('vCMS2');<% End If %>
        cpVal('vCMS3');
        cpVal('vCMS4');
        cpVal('vCMS5');
        cpVal('vCMS6');
        cpVal('vCMS7');
        cpVal('vAktiverad');
      <% End If %>
    }
    
    function local_ResetFields() {
    }
  </script>
  
  <%
  bLockValues = True
  If rsDB.EOF And GetAcc("CMS202") Then bLockValues = False
  %>
  
  <form id="em" method="POST">
  <div class="datablock rect morepadding">
    <div class="legend">Hantera användare</div>
    
    <input type="hidden" id="vID" name="vID" value="<% = cADD_ID %>">
    
    <div class="in_row">
      <div class="text">Titel</div>
      <div class="input">
        <select name="vTitel" <% If Not GetAcc("CMS202") Then Response.Write(" disabled") %>>
          <%
          Set rsTT = Server.CreateObject("ADODB.RecordSet")
          SQL = "SELECT * FROM fsBB_Titlar ORDER By ttSortNr ASC"
          rsTT.Open SQL, Con, 1, 3
          
          Do Until rsTT.EOF
            %>
              <option value="<% = rsTT("ttID") %>" <% If CLng(cADD_Titel) = rsTT("ttID") Then Response.Write(" selected") %>> <% = rsTT("ttForklaring") %> </option>
            <%
          rsTT.MoveNext
          Loop
          
          rsTT.Close
          Set rsTT = Nothing
          %>
        </select>
      </div>
    </div>
    
    <div class="in_row">
      <div class="text">Egen titlel</div>
      <div class="input"><input type="text" class="fill" name="vEgenTitel" maxlength="20" value="<% = cADD_EgenTitel %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Användarnamn</div>
      <div class="input"><input type="text" class="fill notnull" name="vAnvNamn" maxlength="60" value="<% = cADD_AnvNamn %>" <% If bLockValues Then Response.Write(" disabled") %>></div>
    </div>
    
    <div class="in_row">
      <div class="text">E-Postadress</div>
      <div class="input"><input type="text" class="fill notnull" name="vEpost" maxlength="255" value="<% = cADD_Epost %>" <% If Not GetAcc("CMS202") Then Response.Write(" disabled") %>></div>
    </div>
    
    <% If GetAcc("CMS202") Then %>
      <div class="in_line"> </div>
      
      <div class="in_row">
        <div class="text">Nytt lösenord</div>
        <div class="input"><input type="password" class="fill" name="vLosen1" maxlength="50"></div>
      </div>
      
      <div class="in_row">
        <div class="text">Bekräfta lösenord</div>
        <div class="input"><input type="password" class="fill" name="vLosen2" maxlength="50"></div>
      </div>
      
      <div class="in_text"> <p>För att kunna byta lösenord på en användare måste båda dessa fält stämma överrens och lösenordet måste bestå av minst 5 tecken.</p> </div>
    <% End If %>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Namn</div>
      <div class="input"><input type="text" class="fill notnull" name="vNamn" maxlength="50" value="<% = cADD_Namn %>"></div>
    </div>
    
    <div class="in_row">
      <div class="text">Plats</div>
      <div class="input"><input type="text" class="fill" name="vPlats" maxlength="50" value="<% = cADD_Plats %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Hemsida</div>
      <div class="input"><input type="text" class="fill" name="vHemsida" maxlength="255" value="<% = cADD_Hemsida %>"></div>
    </div>
    
    <div class="in_row">
      <div class="text">MSN</div>
      <div class="input"><input type="text" class="fill" name="vMSN" maxlength="255" value="<% = cADD_MSN %>"></div>
    </div>
    
    <div class="in_row">
      <div class="text">ICQ</div>
      <div class="input"><input type="text" class="fill" name="vICQ" maxlength="255" value="<% = cADD_ICQ %>"></div>
    </div>
    
    <div class="in_line"> </div>
    
    <div class="in_row">
      <div class="text">Signatur</div>
      <div class="input"><input type="text" class="fill" name="vSIgnatur" maxlength="255" value="<% = cADD_Signatur %>"></div>
    </div>
    
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
      <textarea id="myText" name="vText"><% = cADD_Profil %></textarea>
    </div>
    
  </div>
  
  <input type="hidden" name="form" value="edit">
  <input type="hidden" name="f" value="<% = sFilter %>">
  <input type="hidden" name="s" value="<% = lPaSida %>">
  <input type="hidden" name="alfa" value="<% = sSendAlfa %>">
  
  <input type="hidden" name="vBannad_cp">
  
  <input type="hidden" name="vCMS_cp">
  
  <input type="hidden" name="vCMS0_cp">
  <input type="hidden" name="vCMS1_cp">
  <input type="hidden" name="vCMS2_cp">
  <input type="hidden" name="vCMS3_cp">
  <input type="hidden" name="vCMS4_cp">
  <input type="hidden" name="vCMS5_cp">
  <input type="hidden" name="vCMS6_cp">
  <input type="hidden" name="vCMS7_cp">
  
  <input type="hidden" name="vAktiverad_cp">
  
  </form>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <% If cADD_HasCMS And Not GetAcc("CMS202") Then bCantSave = True %>
  
  <div class="databox info">
    <div class="inner" style="text-align: center;">
      <input onclick="cpFlds();saveform('em',0);" name="savebtn" class="save" type="button"value="Spara" <% If bCantSave Then Response.Write(" disabled") %>>
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
        saveDate = Trim(CStr(rsDB("aDatumSparad") & " "))
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
  
  <!--
  <div class="databox info">
    <div class="label">Avatar</div>
    <div class="inner" style="text-align: center;">
      <img src="<% = cADD_Avatar_URL %>" style="border: solid 1px #CCC; width: 100px; height: 100px; background-color: #FFF; display: block; margin: 0 auto 5px auto;">
      <input type="button" value="Radera avatar" onclick="if(confirm('Vill du radera avataren? Åtgärden är oåterkallelig!')){}">
    </div>
  </div>
  -->
  
  <% If CLng(cADD_ID) <> CLng(cCMS_ID) Then %>
    <div class="databox info">
      <div class="label">Åtgärder</div>
      <div class="inner" style="text-align: center;">
        <% If GetAcc("CMS202") Then %>
          <div class="chkbox"><input type="checkbox" value="YES" id="vAktiverad" name="vAktiverad" <% If cADD_IsActiv Then Response.Write(" checked") %>></div>
          <div class="chkbox_text">Aktiverad</div>
          <div class="innerseparator"> </div>
        <% End If %>
        <div class="field" style="margin-bottom: 5px;">
          <select id="vBannad">
            <option value="0"> EJ - Bannad </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If cADD_BlockStatus = 1 Then Response.Write(" selected") %>> 1 Vecka - Bannad </option>
            <option value="2" class="levelin" <% If cADD_BlockStatus = 2 Then Response.Write(" selected") %>> 1 Månad - Bannad </option>
            <option value="3" class="levelin" <% If cADD_BlockStatus = 3 Then Response.Write(" selected") %>> 3 Månad - Bannad </option>
            <option value="4" class="levelin" <% If cADD_BlockStatus = 4 Then Response.Write(" selected") %>> 6 Månad - Bannad </option>
            <option value="5" class="levelin" <% If cADD_BlockStatus = 5 Then Response.Write(" selected") %>> 1 År - Bannad </option>
            <option value="6" class="levelin" <% If cADD_BlockStatus = 6 Then Response.Write(" selected") %>> Evigt - Bannad </option>
          </select>
        </div>
        <% If GetAcc("CMS202") Then %>
          <div class="innerseparator"> </div>
          <div class="field">
            <input type="button" value="Radera alla inlägg" onclick="if(confirm('Vill ta bort användarens alla inlägg? Åtgärden är oåterkallelig!')){}" disabled>
          </div>
          <div class="innerseparator"> </div>
          <div class="field">
            <strong>Onlinestatus</strong>
            <% If cADD_Online Then %><p style="color: #0A0; font-weight: bold;">Online</p><% Else %><p style="color: #A00;">Offline</p><% End If %>
          </div>
          <div class="field">
            <input type="button" value="Logga ut" onclick="if(confirm('Vill du logga ut användaren?')){doform('em','logout');}" <% If Not cADD_Online Or cADD_LoggaUt Then %>disabled<% End If %>>
          </div>
        <% End If %>
      </div>
    </div>
  <% Else %>
    <input type="hidden" id="vBannad" value="0">
  <% End If %>
  
  <% If GetAcc("CMS202") Then %>
    <div class="databox info">
      <div class="label">Användarens behörigheter</div>
      <div class="inner">
        <div class="chkbox"><input <% If CLng(cADD_ID) = CLng(cCMS_ID) Then Response.Write(" disabled") %> type="checkbox" value="YES" id="vCMS" name="vCMS" onclick="if(this.checked){showclass('field CMS_BEH');}else{hideclass('field CMS_BEH');}" <% If cADD_HasCMS Then Response.Write(" checked") %>></div>
        <div class="chkbox_text">CMS Inloggning</div>
        
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS0">
            <option value="0"> EJ - CMS Hantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS000") Then Response.Write(" selected") %>> Publicistens ruta </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS1">
            <option value="0"> EJ - Texthantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS100") Then Response.Write(" selected") %>> Lätt texthantering </option>
            <option value="2" class="levelin" <% If HasAcc(cADD_Ratter,"CMS110") Then Response.Write(" selected") %>> Medel texthantering </option>
            <option value="3" class="levelin" <% If HasAcc(cADD_Ratter,"CMS111") Then Response.Write(" selected") %>> Avancerad texthantering </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS2" <% If CLng(cADD_ID) = CLng(cCMS_ID) Then Response.Write(" disabled") %>>
            <option value="0"> EJ - Användarhantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS200") Then Response.Write(" selected") %>> Lätt användarhantering </option>
            <option value="2" class="levelin" <% If HasAcc(cADD_Ratter,"CMS202") Then Response.Write(" selected") %>> Medel användarhantering </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS3">
            <option value="0"> EJ - Forumhantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS300") Then Response.Write(" selected") %>> Lätt forumhantering </option>
            <option value="2" class="levelin" <% If HasAcc(cADD_Ratter,"CMS330") Then Response.Write(" selected") %>> Medel forumhantering </option>
            <option value="3" class="levelin" <% If HasAcc(cADD_Ratter,"CMS333") Then Response.Write(" selected") %>> Avancerad forumhantering </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS4">
            <option value="0"> EJ - Databashantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS400") Then Response.Write(" selected") %>> Lätt databashantering </option>
            <option value="2" class="levelin" <% If HasAcc(cADD_Ratter,"CMS440") Then Response.Write(" selected") %>> Medel databashantering </option>
            <option value="3" class="levelin" <% If HasAcc(cADD_Ratter,"CMS444") Then Response.Write(" selected") %>> Avancerad databashantering </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS5">
            <option value="0"> EJ - Statiskt material </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS500") Then Response.Write(" selected") %>> Statiskt material </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS6">
            <option value="0"> EJ - Mediahantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS600") Then Response.Write(" selected") %>> Mediahantering </option>
          </select>
        </div>
        <div class="field CMS_BEH" style="display: none;">
          <select id="vCMS7">
            <option value="0"> EJ - Annonshantering </option>
            <option class="separator" disabled> &nbsp; </option>
            <option value="1" class="levelin" <% If HasAcc(cADD_Ratter,"CMS700") Then Response.Write(" selected") %>> Annonshantering </option>
          </select>
        </div>
        
        <script type="text/javascript">
          if(document.getElementById("vCMS").checked){showclass('field CMS_BEH');}else{hideclass('field CMS_BEH');}
        </script>
      </div>
    </div>
  <% End If %>
  
  <div class="databox info">
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