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
If Not GetAcc("CMS2") Then Response.Redirect("/")
%>

<%
cADD_ID = mGet("e","123",0)

' #### BEHÖRIGHET ####
  If NOT GetAcc("CMS202") Then sBFilter = ""
' ####################

Con_Open
Set rsDB = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT * FROM fsBB_Anv LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID WHERE aID = " & CLng(cADD_ID) & sBFilter
rsDB.Open SQL, Con
%>

<% If rsDB.EOF Then %>
  <div class="minimess">Anv&auml;ndaren &auml;r &auml;nnu inte sparad.</div>
<% Else %>
  <% If rsDB("aBlockadTill") >= Now Then %>
    <div class="field"><strong>Bannad tom:</strong><br><p><% = rsDB("aBlockadTill") %></p></div>
    <div class="innerseparator"> </div>
  <% ENd If %>

  <div class="field"><strong>Blev medlem:</strong><br><p><% = rsDB("aMedlemsedan") %></p></div>
  <div class="field"><strong>Senast inloggad:</strong><br><p><% = rsDB("aInloggadSenast") %></p></div>
  
  <div class="innerseparator"> </div>
  
  <div class="field"><strong>IP (Reg):</strong><br><p><% = rsDB("aIn_IP_Reg") %></p></div>
  <div class="field"><strong>IP (Login):</strong><br><p><% = rsDB("aIn_IP_Login") %></p></div>
  <div class="field"><strong>IP (Fel):</strong><br><p><% = rsDB("aIn_IP_Failed") %></p></div>
<% End If %>

<%
rsDB.Close
Set rsDB = Nothing
Con_Close
%>