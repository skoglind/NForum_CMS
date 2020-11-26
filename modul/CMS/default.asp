<% 
  cON_PAGE = "Sammanfattning - CMS"
%>

<!--#INCLUDE FILE="../../_deftop.asp"-->
  
  <%
  Con_Open
  Set rsDB = Server.CreateObject("ADODB.Recordset")
  SQL = "SELECT * FROM cms_InfoBlock ORDER BY ifSortNr ASC"
  rsDB.Open SQL, Con
  %>
  
  <div class="datablock rect">
    <% bInfoBlock = GetAcc("CMS0") %>
    
    <div class="legend">Information</div>
    <% If bInfoBlock Then %><div class="itextedit"><input type="button" value="Ny text" onclick="location.href='_edit.asp';"></div><% End If %>
    
    <% If rsDB.EOF Then %>
      <div class="itext">
        <p style="text-align: center;"> <br><em>Det finns ingen information just nu.</em><br><br> </p>
      </div>
    <% End If %>
    
    <% Do Until rsDB.EOF %>
      <div class="itext">
        <p><strong><% = sEncode(rsDB("ifTitel")) %></strong> <% = Replace(sEncode(rsDB("ifTextM")), Chr(13), "<br>") %> </p>
      </div>
      <% If bInfoBlock Then %><div class="itextedit"><input type="button" value="Ändra" onclick="location.href='_edit.asp?e=<% = rsDB("ifID") %>';"> <input type="button" style="color: #A00; font-weight: bold;" value="X" onclick="if(confirm('Vill du ta bort texten?')){location.href='__do.asp?e=<% = rsDB("ifID") %>&amp;a=del';}"></div><% End If %>
      <% rsDB.MoveNext %>
    <% Loop %>
  </div>
  
  <%
  rsDB.Close
  Set rsDB = Nothing
  Con_Close
  %>
  
  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->
  
  <% If GetAcc("CMS11") Then %>
    <div class="databox info">
      <div class="label">Inväntar publicering</div>
      <div class="inner">
        <%
        Con_Open
          lStatus_0 = Con.ExeCute("SELECT COUNT(*) FROM cms_Nyheter WHERE nStatus = 2")(0)
          lStatus_1 = Con.ExeCute("SELECT COUNT(*) FROM cms_Recensioner WHERE rStatus = 2")(0)
          lStatus_2 = Con.ExeCute("SELECT COUNT(*) FROM cms_Artiklar WHERE aaStatus = 2")(0)
          
          lStatus_3 = Con.ExeCute("SELECT COUNT(*) FROM cms_Speltrix WHERE xStatus = 2")(0)
          lStatus_4 = Con.ExeCute("SELECT COUNT(*) FROM cms_Konsoltrix WHERE xStatus = 2")(0)
        Con_Close
        %>
        <table class="list" cellpadding=0 cellspacing=0>
          <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> nyheter </td></tr>
          <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_1 %> </td><td class="td3"> recensioner </td></tr>
          <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_2 %> </td><td class="td3"> artiklar </td></tr>
        </table>
        <div class="innerseparator"> </div>
        <table class="list" cellpadding=0 cellspacing=0>
          <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_3 %> </td><td class="td3"> trix (spel) </td></tr>
          <tr><td class="td1"> <img src="/design/icons/papper_2_sm.png"> </td><td clasS="td2"> <% = lStatus_4 %> </td><td class="td3"> trix (konsol) </td></tr>
        </table>
      </div>
    </div>
  <% End If %>
  
  <div class="databox info">
    <div class="label">Statistik</div>
    <div class="inner">
      <%
      Con_Open
        lStatus_0 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aAktiverad = 1 AND aBlockadTill < '" & Now & "'")(0)
        lStatus_1 = Con.ExeCute("SELECT COUNT(*) FROM cms_Nyheter WHERE nStatus = 4")(0)
        lStatus_2 = Con.ExeCute("SELECT COUNT(*) FROM cms_Recensioner WHERE rStatus = 4")(0)
        lStatus_3 = Con.ExeCute("SELECT COUNT(*) FROM cms_Artiklar WHERE aaStatus = 4")(0)
        lStatus_4 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tForum <> 32 AND tStatus_Trad = 1")(0)
        lStatus_5 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tForum <> 32 AND tStatus_Trad = 0")(0)
        lStatus_6 = Con.ExeCute("SELECT COUNT(*) FROM fsBB_PM WHERE pRaderadFran = 0 OR pRaderadTill = 0")(0)
      Con_Close
      %>
      <% If GetAcc("CMS2") Then %>
        <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
          <tr><td class="td1"> <img src="/design/icons/user_0_sm.png"> </td><td clasS="td2"> <% = lStatus_0 %> </td><td class="td3"> användare </td></tr>
        </table>
        <div class="innerseparator"> </div>
      <% End If %>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <tr><td class="td1"> <img src="/design/icons/papper_4_sm.png"> </td><td clasS="td2"> <% = lStatus_1 %> </td><td class="td3"> nyheter </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_4_sm.png"> </td><td clasS="td2"> <% = lStatus_2 %> </td><td class="td3"> recensioner </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_4_sm.png"> </td><td clasS="td2"> <% = lStatus_3 %> </td><td class="td3"> artiklar </td></tr>
      </table>
        <div class="innerseparator"> </div>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lStatus_4 %> </td><td class="td3"> forumtrådar </td></tr>
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lStatus_5 %> </td><td class="td3"> foruminlägg </td></tr>
      </table>
        <div class="innerseparator"> </div>
      <table class="list" cellpadding=0 cellspacing=0 style="float: left;">
        <tr><td class="td1"> <img src="/design/icons/papper_1_sm.png"> </td><td clasS="td2"> <% = lStatus_6 %> </td><td class="td3"> PM </td></tr>
      </table>
    </div>
  </div>
      
<!--#INCLUDE FILE="../../_defbottom.asp"-->     