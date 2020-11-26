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

<% If GetAcc("CMS4") Then %>
  <%
  q = LCase(Trim(MakeLegal_Large(Request.QueryString("t"))))
  
  k = Request.QueryString("k")
  If Not IsNumeric(k) Or k = Empty Then k = 0
  k = CLng(k)
  
  s = Request.QueryString("s")
  If Not IsNumeric(s) Or s = Empty Then s = 0
  s = CLng(s)
  
  e = Trim(MakeLegal(Request.QueryString("e")))
  If Not IsNumeric(e) Or e = Empty Then 
    transID = 0
    exklID  = Trim(MakeLegal(e))
  Else
    transID = CLng(e)
    exklID  = "RR"
  End If
  
  Con_Open
  Set rsQ = Server.CreateObject("ADODB.RecordSet")
  SQL = "SELECT TOP 10 * FROM cms_Speltitlar RIGHT JOIN cms_Spel ON cms_Spel.sID = cms_Speltitlar.tSpelID WHERE tTitel LIKE '%" & q & "%' AND sKonsol = " & k & " AND tSpelID <> " & s & " AND tID <> " & transID & " ORDER BY tTitel ASC"
  rsQ.Open SQL, Con
  
  allCnt = Con.ExeCute("SELECT COUNT(*) FROM cms_Speltitlar RIGHT JOIN cms_Spel ON cms_Spel.sID = cms_Speltitlar.tSpelID WHERE tTitel LIKE '%" & q & "%' AND sKonsol = " & k & " AND tSpelID <> " & s & " AND tID <> " & transID)(0)
  If Not IsNumeric(allCnt) Or allCnt = Empty Then allCnt = 0
  %>
  
  <% If rsQ.EOF Then %>
    &nbsp;
  <% Else %>
    <div class="holder">
      <p>F&ouml;ljande titlar finns redan i databasen med liknande namn,</p>
      <ul>
        <% For zx =  1 To 10 %>
          <% If rsQ.EOF Then Exit For %>
            <% allCnt = allCnt - 1 %>
            <li> <span class="region"><% = lstRegion(rsQ("tRegion")) %></span> <% = sEncode(rsQ("tTitel")) %> </li>
          <% rsQ.MoveNext %>
        <% Next %>
      </ul>
      <p><% If allCnt > 0 Then %>... och ytterligare <% = allCnt %> titlar till.<% End If %></p>
    </div>
  <% End If %>
  
  <%
  rsQ.Close
  Set rsQ = Nothing
  Con_Close
  %>

<% Else %>
  <div class="holder">
    There's no data for you, my friend!
  </div>
<% End If %>