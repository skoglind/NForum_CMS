<% If Session.Value("CMS_LOGIN") Then Response.Redirect("/") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<!--#INCLUDE FILE="cms_Config.asp"-->
<!--#INCLUDE FILE="cms_Functions.asp"-->

<%
Session.Value("LOGINCODE") = Make_LC()
%>

<html>
  <head>
    <title> Logga in | <% = CMS_SITENAME %>  </title>
    <meta http-equiv="content-type" content="text/html; CHARSET=ISO-8859-1">
    <meta http-equiv="content-language" content="sv">
    <meta name="ROBOTS" content="NOINDEX, NOFOLLOW">
    <link rel="stylesheet" type="text/css" href="res/standard.css">
  </head>
  <body onload="document.getElementById('anamn').focus();">
  
    <div class="box_type1">
      <div class="inner" style="background-color: #FFF;">
        <form action="_processlogin.asp" method="POST">
          <fieldset>
            <legend> Logga in [<% = CMS_SITENAME %>] </legend>
              <div class="row"><div>Användarnamn</div><input name="cms_anvnamn" id="anamn" type="text" maxlength=50></div>
              <div class="row"><div>Lösenord</div> <input name="cms_passwd" id="passwd" type="password" maxlength=50></div>
              <div class="button"><input type="submit" value="Logga in"></div>
              
              <input type="hidden" name="lc" value="<% = Session.Value("LOGINCODE") %>">
          </fieldset>
        </form>
      </div>
    </div>
    
    <% If Session.Value("ERR_Code") > 0 Or CMS_HALT Then %>
      <div class="box_type3">
        <div class="inner">
          <form><fieldset class="warning">
            <strong>Fel vid inloggning!</strong><br>
            <% If CMS_HALT Then %>
              Systemet nerstängt! Alla funktioner avstängda! Kontakta din administratör om varför!
            <% End If %>
              
            <% Select Case Session.Value("ERR_Code") %>
            <% Case 1 %>
              Inloggningen misslyckades. Användarnamnet och/eller lösenordet var felaktigt.
            <% Case 2 %>
              För många felaktiga inloggningsförsök, du är låst från att logga in under <strong><% = LOGIN_MINUTER %></strong> minuter.
            <% Case 3 %>
              Du måste accpetera cookies för att kunna logga in.
            <% Case 4 %>
              Du måste skicka med en referer för att kunna logga in.
            <% Case 5 %>
              Login är inte aktiverat, kontakta din administratör.
            <% End Select %>
            <% Session.Value("ERR_Code") = 0 %>
          </fieldset></form>
        </div>
      </div>
    <% End If %>
    
    <!--[if IE 7]>
      <div class="box_type3">
        <div class="inner">
          <form><fieldset class="low_warning">
            <strong>OBS!</strong><br>
            Det rekommenderas starkt att du hämtar hem Firefox om du inte redan använder den webbläsaren, <a href="http://www.getfirefox.com" target="_blank">hämta Mozilla Firefox</a>.
          </fieldset></form>
        </div>
      </div>
    <![endif]-->
    
    <!--[if lt IE 7]>
      <div class="box_type3">
        <div class="inner">
          <form><fieldset class="warning">
            <strong>OBS!</strong><br>
            Endast optimerat för IE7+ och Firefox2+ och det är därför inte säkert att sidan funkar och ser ut som den ska i andra webbläsare.<br><br>
            Det rekommenderas starkt att du hämtar hem Firefox om du inte redan använder den webbläsaren, <a href="http://www.getfirefox.com" target="_blank">hämta Mozilla Firefox</a>.
          </fieldset></form>
        </div>
      </div>
    <![endif]-->
  
  </body>
</html>