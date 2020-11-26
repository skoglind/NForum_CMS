<% 
  cON_PAGE = "Meddelande - CMS"
  
  sM_Meddelande = Session.value("PBM_Message")
  sM_Lank       = Session.value("PBM_Lank")
  
  Session.value("PBM_Message")  = ""
  Session.value("PBM_Lank")     = ""
%>

<!--#INCLUDE FILE="_deftop.asp"-->

  <div class="datablock rect">
    <div class="legend">Meddelande</div>
    <div class="textblock">
      <% If sM_Meddelande = Empty Then %>
        <h2>Inget meddelande!</h2>
        <p>Det finns inget meddelande att visa just nu.</p>
      <% Else %>
        <% = sM_Meddelande %>
        <p style="text-align: right;"><a href="<% = sM_Lank %>">Fortsätt »</a></p>
      <% End If %>
    </div>
  </div>

  <!-- ## DELIMITER ## --></div><div class="extra"><!-- ## DELIMITER ## -->

<!--#INCLUDE FILE="_defbottom.asp"-->     