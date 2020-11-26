<% If Not Session.Value("CMS_LOGIN") Then Response.Redirect("/login.asp") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<%
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.CacheControl = "no-cache"
%>

<!--#INCLUDE FILE="../cms_Config.asp"-->
<!--#INCLUDE FILE="../cms_Constant.asp"-->
<!--#INCLUDE FILE="../cms_Functions.asp"-->
<!--#INCLUDE FILE="../cms_Lists.asp"-->

<% If CMS_HALT Then Response.Write("Systemet nerstängt! Alla funktioner avstängda! Kontakta din administratör om varför!") : Response.End %>

<html>
  <head>
    <title> Välj konsol... | <% = CMS_SITENAME %>  </title>
    <meta http-equiv="content-type" content="text/html; CHARSET=ISO-8859-1">
    <meta http-equiv="content-language" content="sv">
    <link rel="stylesheet" type="text/css" href="/res/picker.css">
    <!--[if lt IE 7.]><script defer type="text/javascript" src="pngfix.js"></script><![endif]-->
    <script type="text/javascript" src="/res/ajax.js"></script>
    <script type="text/javascript" src="/res/standard.js"></script>
  </head>
  <body onload="document.getElementById('searchval').focus();">
  
    <div class="holder">
      <div class="label">
        <h1>Välj konsol...</h1>
      </div>
    
      <%
      q = Trim(Request.QueryString("q") & " ")
      If Len(q) > 0 Then q = MakeLegal_Large(q)
      %>
      
      <div class="box">
        <div class="inner">
          <form method="get">
            <input type="text" name="q" style="width: 220px;" value="<% = q %>" id="searchval">
            <input type="submit" style="width: 60px;" value="Sök">
          </form>
        </div>
      </div>
      
      <%
      If Len(q) > 2 Then
        SQL = "SELECT * FROM cms_KonsolTitlar LEFT JOIN cms_Konsol ON kID = tKonsolID WHERE tTitel LIKE '%" & CStr(q) & "%' ORDER BY tTitel ASC"
      Else
        SQL = "SELECT * FROM cms_KonsolTitlar LEFT JOIN cms_Konsol ON kID = tKonsolID ORDER BY tTitel ASC"
      End If
      
      Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      rsDB.Open SQL, Con
      
      ' #### PAGING ####
      lMaxPosterPerSida = 10
      lAntalPoster = Con.ExeCute("SELECT COUNT(*) FROM cms_KonsolTitlar")(0)
      If Not IsNumeric(lAntalPoster) Or lAntalPoster = Empty Then lAntalPoster = 0
      lAntalSidor = CLng(RoundUp(lAntalPoster, lMaxPosterPerSida))
      
      lPaSida = Request.QueryString("s")
      If Not IsNumeric(lPaSida) Or lPaSida = Empty Then lPaSida = 1
      lPaSida = CLng(lPaSida)
      If lPaSida < 1 Then lPaSida = 1
      If lPaSida > lAntalSidor Then lPaSida = lAntalSidor
      
      If Not lPaSida = 1 And lAntalPoster > 0 Then rsDB.Move (lPaSida - 1) * lMaxPosterPerSida
      ' ################
      %>
      
      <div class="box" style="background-color: #EAEAEA; height: 240px;">
        <div class="inner">
          <% If rsDB.EOF Then %>
            <p><em>Inga konsoler kunde listas.</em></p>
            <% noSel = True %>
          <% Else %>
            <form>
              <ul>
                <% For zx = 1 To 10 %>
                  <% If rsDB.EOF Then Exit For %>
                  <li> <input type="radio" name="spel" value="<% = rsDB("tKonsolID") %>" <% If zx = 1 Then Response.Write(" checked") %>> <div class="text"><% = lstKonsolXShort(rsDB("kKonsol")) & " | " & sEncode(rsDB("tTitel")) %></div> <input type="hidden" name="ibox<% = rsDB("tKonsolID") %>" id="ibox<% = rsDB("tKonsolID") %>" value="<% = lstKonsolXShort(rsDB("kKonsol")) & " | " & sEncode(rsDB("tTitel")) %>"> </li>
                  <% rsDB.MoveNext %>
                <% Next %>
              </ul>
            </form>
            <% noSel = False %>
          <% End If %>
        </div>
      </div>
      
      <%
      rsDB.Close
      Set rsDB = Nothing
      Con_Close
      %>
      
      <div class="move">
        <div class="inner" style="width: 290px;">
          <% If lPaSida >= lAntalSidor Then %>
            <a href="#" style="float: right; color: #AAA;">Nästa »</a>
          <% Else %>
            <a href="?s=<% = lPaSida + 1 %>&q=<% = q %>" style="float: right;">Nästa »</a>
          <% End If %>
          
          <% If lPaSida < 2 Then %>
            <a href="#" style="float: left; color: #AAA;">« Föregående</a>
          <% Else %>
            <a href="?s=<% = lPaSida - 1 %>&q=<% = q %>" style="float: left;">« Föregående</a>
          <% End If %>
        </div>
      </div>
      
      <div class="buttons">
        <div class="inner">
          <input type="button" value="Stäng" style="float: right;" onclick="window.close();">
          <input type="button" value="OK" style="float: left; font-weight: bold;" onclick="sendbackdata(getSelectedRadio('spel'));" <% If noSel Then Response.Write(" disabled") %>>
        </div>
      </div>
    </div>
    
  </body>
</html>