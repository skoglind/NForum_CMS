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
If Not GetAcc("CMS1") Then Response.Redirect("/")
%>

<%
cADD_ID = mGet("e","123",0)

' #### BEHÖRIGHET ####
  If NOT GetAcc("CMS111") Then sBFilter = " AND aaSkapadAv = " & cCMS_ID & " AND NOT aaStatus = 0"
' ####################

Con_Open
Set rsDB = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT *, anvDB1.aNamn AS Anv1, anvDB2.aNamn AS Anv2 " & _
      "FROM (cms_Artiklar " & _ 
      "LEFT JOIN fsBB_Anv AS AnvDB1 ON cms_Artiklar.aaSkapadAv = AnvDB1.aID) " & _
      "LEFT JOIN fsBB_Anv AS AnvDB2 ON cms_Artiklar.aaPubliceradAv = AnvDB2.aID " & _
      "WHERE aaID = " & CLng(cADD_ID) & sBFilter
rsDB.Open SQL, Con
%>

<% If rsDB.EOF Then %>
  <div class="minimess">Artikeln &auml;r &auml;nnu inte sparad.</div>
<% Else %>
  <% lStatus = rsDB("aaStatus") %>

  <% If GetAcc("CMS11") Then %>
    <div class="field">
      <select id="vStatus" name="vStatus" onchange="if(this.value=='4'){document.getElementById('vPublTyp').disabled=false;}else{document.getElementById('vPublTyp').disabled=true;document.getElementById('vPublDatum').disabled=true;}">
        <option value="A"> &Auml;ndra status... </option>
        <option class="separator" disabled> &nbsp; </option>
        <option value="4" class="levelin"> Publicerad </option>
        <option value="3" class="levelin"> Publicering nekad </option>
        <option value="2" class="levelin"> Inv&auml;ntar publicering </option>
        <option value="1" class="levelin"> Under bearbetning </option>
      </select>
    </div>
    
    <div class="field">
      <select id="vPublTyp" name="vPublTyp" disabled onchange="if(this.value=='4'){document.getElementById('vPublDatum').disabled=false;}else{document.getElementById('vPublDatum').disabled=true;}">
        <option value="A"> Typ av publicering... </option>
        <option class="separator" disabled> &nbsp; </option>
        <option value="1" class="levelin"> Publicera direkt </option>
        <option value="2" class="levelin"> Publicera imorgon </option>
        <option value="3" class="levelin"> Publicera n&auml;sta vardag </option>
        <option value="4" class="levelin"> Publicera p&aring; datum som &auml;r valt nedan </option>
      </select>
    </div>
    
    <div class="field">
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
  <div class="field"><strong>Skapad:</strong><br><p><% = rsDB("aaDatumSkapad") %></p></div>
  <div class="field"><strong>Skapad av:</strong><br><p><% = Server.HTMLEncode(rsDB("Anv1")) %></p></div>
  
  <% If GetAcc("CMS111") Then %>
    <div class="innerseparator"> </div>
  
    <div class="field">
      <select id="vNySkapare" name="vNySkapare">
        <option value="0"> &Auml;ndra skapare... </option>
        <option class="separator" disabled> &nbsp; </option>
        <% Call ListCMSUsers() %>
      </select>
    </div>
  <% End If %>
  
  <% If lStatus = 4 Then %>
    <div class="innerseparator"> </div>
    <% If rsDB("aaDatumPublicerad") > Now Then %>
      <div class="field"><strong>Kommer publiceras:</strong><br>
    <% Else %>
      <div class="field"><strong>Publicerad:</strong><br>
    <% End If %>
    <p><% = FormatDateTime(rsDB("aaDatumPublicerad"), vbShortDate) %> (<% = Server.HTMLEncode(WeekDayName(weekday(rsDB("aaDatumPublicerad")))) %>)</p></div>
    <div class="field"><strong>Publicerad av:</strong><br><p><% = Server.HTMLEncode(rsDB("Anv2")) %></p></div>
  <% End If %>
<% End If %>

<%
rsDB.Close
Set rsDB = Nothing
Con_Close
%>

<div class="innerseparator"> </div>
<div class="field"><strong>Flash:</strong> 301x200px<br>

<%
Con_Open
 Set rsTN = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Artiklar LEFT JOIN cms_Bild ON cms_Artiklar.aaFlash = cms_Bild.bID WHERE aaID = " & CLng(cADD_ID)
    rsTN.Open SQL, Con
    
      anyHit = True
      If rsTN.EOF Then
        anyHit = False
        cADD_ID = 0
      Else
        cADD_ID = rsTN("aaID")
        If IsNull(rsTN("bID")) Then 
          anyHit = False
        Else
          bildID  = CLng(rsTN("bID"))
          bildTyp = Trim(rsTN("bTyp"))
        End If
      End IF
    
    rsTN.Close
  Set rsTN = Nothing
  
  If CLng(cADD_ID) = 0 Then
    bildID = GetImgIDByExclusive()
    If CLng(bildID) > 0 Then anyHit = True
  End If
Con_Close
%>

<% If GetAcc("CMS111") Then %>
<form id="smallupload" method="POST" enctype="multipart/form-data">
  <div class="field" style="text-align: center;"><input type="file" name="thaFile" size="16"></div>
  <div class="field" style="text-align: center;"><input type="button" value="Ladda upp" onclick="smalluploadimg('smallupload');"></div>
  <input type="hidden" value=<% = cADD_ID %> name="lID">
  <input type="hidden" value="artiklar" name="sArea">
</form>
<% End If %>

<% If anyHit Then %>
  <%
  Con_Open
    sOriginal = ImgOriginal(bildID)
    If sOriginal = "NO_IMG" Then
      anyPic = False
    Else
      anyPic = True
    End IF
  Con_Close
  
  sFilNamn = "/cms_Img.asp?e=" & bildID & "&w=150&h=150"
  %>
  
  <div class="innerseparator"> </div>
  <img src="<% If anyPic Then Response.Write(sFilNamn) %>" style="float: left; border: solid 1px #CCC; width: 150px; height: 150px; background-color: #FFF; display: block; margin: 0 12px 5px 12px;">
  <% If GetAcc("CMS111") Then %><div class="field" style="text-align: center;"><input type="button" value="Radera bild" onclick="if(confirm('Vill du radera bilden?')){smalldeleteimg(<% = cADD_ID %>,'artiklar');}"></div><% End If %>
<% End If %>