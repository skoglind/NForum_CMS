<% If Not Session.Value("CMS_LOGIN") Then Response.Redirect("/login.asp") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<%
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.CacheControl = "no-cache"
%>

<!--#INCLUDE FILE="cms_Config.asp"-->
<!--#INCLUDE FILE="cms_Constant.asp"-->
<!--#INCLUDE FILE="cms_Functions.asp"-->
<!--#INCLUDE FILE="cms_Lists.asp"-->

<% If CMS_HALT Then Response.Write("Systemet nerstängt! Alla funktioner avstängda! Kontakta din administratör om varför!") : Response.End %>

<%
Session.Value("LOGINCODE") = Make_LC()
%>

<html>
  <head>
    <title> <% = cON_PAGE %> | <% = CMS_SITENAME %>  </title>
    <meta http-equiv="content-type" content="text/html; CHARSET=ISO-8859-1">
    <meta http-equiv="content-language" content="sv">
    <link rel="stylesheet" type="text/css" href="/res/standard.css">
    <!--[if IE 7]><link rel="stylesheet" type="text/css" href="/res/ie7.css"><![endif]-->
    <!--[if IE 6]><link rel="stylesheet" type="text/css" href="/res/ie6.css"><![endif]-->
    <!--[if lt IE 7]><script defer type="text/javascript" src="pngfix.js"></script><![endif]-->
    <script type="text/javascript" src="/res/ajax.js"></script>
    <script type="text/javascript" src="/res/standard.js"></script>
  </head>
  <body class="system" onload="StayOnline();">
  
    <div id="OUTER"><div id="INNER">
  
    <div class="box_type2">
      <div class="closeborder">
        <div class="topborder">
          <p style="float: right; text-align: right;">
            <input type="button" value="Logga ut" onclick="if(confirm('Vill du logga ut?')){location.href=('/_processlogout.asp');}">
          </p>
          <p style="float: left;" class="title">› <% = CMS_SITENAME %> Administration</p>
        </div>
      </div>
    
      <div class="menu">
        <ul>       
          <li class="title"> <span>CMS</span> </li>
          <ul>
            <li> <a href="/Modul/CMS/">Sammanfattning</a> </li>
            <li> <a href="/Modul/CMS/ReadRSS/_show.asp">RSS-Strömmar</a> </li>
            <li> <a class="dis">Panik!!!</a> </li>
            <!-- <li> <a class="dis">Diskussion</a> </li> -->
            <li> <a href="/Modul/CMS/Omrostningar/_show.asp">Omröstningar</a> </li>
            <!-- <li> <a class="dis">Hjälp</a> </li> -->
          </ul>
        </ul>
        
        <ul>
          <li class="title"> <span>System</span> </li>
          <ul>
            <% If GetAcc("CMS2") Then %><li> <a href="/Modul/fsBB/Anvandare/_show.asp">Användare</a> </li><% End If %>
            <% If GetAcc("CMS202") Then %><li> <a href="/Modul/fsBB/Titlar/_show.asp">Titlar</a> </li><% End If %>
            <!-- <% If GetAcc("CMS33") Then %><li> <a class="dis">Mass-PM</a> </li><% End If %> -->
            <!-- <% If GetAcc("CMS202") Then %><li> <a class="dis">Mass-Mail</a> </li><% End If %> -->
            <!-- <li> <a class="dis">Mediabibliotek</a> </li> -->
          </ul>
        </ul>
        
          <% If GetAcc("CMS3") Then %>
            <ul>
              <li class="title"> <span>Forum</span> </li>
              <ul>
                <!-- <li> <a class="dis">Statistik</a> </li> -->
                <% If GetAcc("CMS333") Then %><li> <a href="/Modul/fsBB/Forum/_show.asp">Forumkategorier</a> </li><% End If %>
                <% If GetAcc("CMS3") Then %><li> <a href="/Modul/fsBB/Anmalningar/_show.asp">Anmälningar</a> </li><% End If %>
              </ul>
            </ul>
          <% End If %>
          <% If GetAcc("CMS1") Then %>
            <ul>
              <li class="title"> <span>Texter</span> </li>
              <ul>
                <li> <a href="/Modul/Texter/Nyheter/_show.asp">Nyheter</a> </li>
                <li> <a href="/Modul/Texter/Recensioner/_show.asp">Recensioner</a> </li>
                <li> <a href="/Modul/Texter/Artiklar/_show.asp">Artiklar</a> </li>
                <!-- <li> <a class="dis">Guider</a> </li> -->
                <% If GetAcc("CMS111") Then %><li> <a href="/Modul/Texter/Trix/_show.asp">Trix (Spel)</a> </li><% End If %>
                <% If GetAcc("CMS111") Then %><li> <a href="/Modul/Texter/Konsoltrix/_show.asp">Trix (Konsol)</a>  </li><% End If %>
              </ul>
            </ul>
          <% End If %>
          <% If GetAcc("CMS4") Then %>
            <ul>
              <li class="title"> <span>Databas</span> </li>
              <ul>
                <!-- <li> <a href="/Modul/Databas/Bidrag/_show.asp">Bidrag</a> </li> -->
                <li> <a href="/Modul/Databas/Spel/_show.asp">Spel</a> </li>
                <li> <a href="/Modul/Databas/Spelserier/_show.asp">Spelgrupper</a> </li>
                <li> <a href="/Modul/Databas/Foretag/_show.asp">Företag</a> </li>
                <li> <a href="/Modul/Databas/Konsol/_show.asp">Konsoler</a> </li>
                <li> <a href="/Modul/Databas/Tillbehor/_show.asp">Tillbehör</a> </li>
                <!-- <li> <a class="dis">Regioner</a> </li> -->
                <!-- <li> <a class="dis">Persongalleri</a> </li> -->
              </ul>
            </ul>
            
            <!--
            <ul>
              <li class="title"> <span>Snabblistning av Spel</span> </li>
              <ul>
                <li> <a href="/Modul/Databas/Q_List/_work.asp">Från lista</a> </li>
                <li> <a href="/Modul/Databas/Q_OaO/_work.asp">En och en</a> </li>
                <li> <a href="/Modul/Databas/Q_Bidrag/_show.asp">Från bidrag</a> </li>
              </ul>
            </ul>
            -->
          <% End If %>
          <% If GetAcc("CMS5") Then %>
          <!-- 
          <ul>
            <li class="title"> <span>Statiskt material</span> </li>
            <ul>
              <li> <a class="dis">Info / Information</a> </li>
              <li> <a class="dis">Info / F.A.Q</a> </li>
            </ul>
          </ul>
          -->
          <% End If %>
        </ul>
          
      </div>
      
      <div class="data">