<%
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.CacheControl = "no-cache"
%>

<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->

<%
If Not GetAcc("CMS111") Then Response.Redirect("/")
%>

<%
cADD_ID = mGet("e","123",0)

' #### BEHÖRIGHET ####
  'If NOT GetAcc("CMS111") Then sBFilter = " AND rSkapadAv = " & cCMS_ID & " AND NOT rStatus = 0"
' ####################

Con_Open
Set rsDB = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT *, anvDB1.aNamn AS Anv1, anvDB2.aNamn AS Anv2 " & _
      "FROM (cms_Konsoltrix " & _ 
      "LEFT JOIN fsBB_Anv AS AnvDB1 ON cms_Konsoltrix.xSkapadAv = AnvDB1.aID) " & _
      "LEFT JOIN fsBB_Anv AS AnvDB2 ON cms_Konsoltrix.xPubliceradAv = AnvDB2.aID " & _
      "WHERE xID = " & CLng(cADD_ID) & sBFilter
rsDB.Open SQL, Con
%>

<% If rsDB.EOF Then %>
  <div class="minimess">Tips & Trixet &auml;r &auml;nnu inte sparat.</div>
<% Else %>
  <% lStatus = rsDB("xStatus") %>

  <% If GetAcc("CMS11") Then %>
    <div class="field">
      <select id="vStatus" name="vStatus" onchange="if(this.value=='4'){document.getElementById('vPublTyp').disabled=false;}else{document.getElementById('vPublTyp').disabled=true;document.getElementById('vPublDatum').disabled=true;}">
        <option value="A"> &Auml;ndra status... </option>
        <option class="separator" disabled> &nbsp; </option>
        <option value="4" class="levelin"> Publicerad </option>
        <option value="3" class="levelin"> Publicering nekad </option>
        <option value="2" class="levelin"> Inv&auml;ntar publicering </option>
      </select>
    </div>
    
    <div class="field">
      <select id="vPublTyp" name="vPublTyp" disabled onchange="if(this.value=='4'){document.getElementById('vPublDatum').disabled=false;}else{document.getElementById('vPublDatum').disabled=true;}">
        <option value="A"> Typ av publicering... </option>
        <option class="separator" disabled> &nbsp; </option>
        <option value="1" class="levelin"> Publicera direkt </option>
      </select>
    </div>
    
    <div class="field" style="display: none;">
      <select id="vPublDatum" name="vPublDatum" disabled>
        <option value="A"> Datum... </option>
        <option class="separator" disabled> &nbsp; </option>
        <% For zx = 1 To 60 %>
          <% nVal = DateAdd("d", zx, Date) %>
          <% nDayName = WeekDayName(weekday(nVal)) %>
          <option value="<% = nVal %>" class="levelin" <% If Left(nDayName, "1") = "s" Then Response.Write(" style='background-color: #fdcece;'") %>> <% = nVal %> (<% = Server.HTMLEncode(nDayName) %>) </option>
        <% Next %>
      </select>
    </div>
    
    <div class="innerseparator"> </div>
  <% End If %>
  
  <div class="field"><strong>Status:</strong><br><p><% = Server.HTMLEncode(lstStatus(lStatus)) %></p></div>
  <div class="field"><strong>Skapad:</strong><br><p><% = rsDB("xDatumSkapad") %></p></div>
  <div class="field"><strong>Skapad av:</strong><br><p><% = Server.HTMLEncode(rsDB("Anv1")) %></p></div>
  
  <% If lStatus = 4 Then %>
    <div class="innerseparator"> </div>
    <div class="field"><strong>Publicerad:</strong><br>
    <p><% = FormatDateTime(rsDB("xDatumPublicerad"), vbShortDate) %> (<% = Server.HTMLEncode(WeekDayName(weekday(rsDB("xDatumPublicerad")))) %>)</p></div>
    <div class="field"><strong>Publicerad av:</strong><br><p><% = Server.HTMLEncode(rsDB("Anv2")) %></p></div>
  <% End If %>
<% End If %>

<%
rsDB.Close
Set rsDB = Nothing
Con_Close
%>