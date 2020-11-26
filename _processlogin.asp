<!--#INCLUDE FILE="cms_Config.asp"-->
<!--#INCLUDE FILE="cms_Functions.asp"-->

<% If CMS_HALT Then Response.Redirect("/login.asp") %>

<%
If Request.Form("lc") <> Session.Value("LOGINCODE") AND Not Request.Form("lc") = Empty Then
  Session.Value("LOGINCODE") = Make_LC()
  Session.Value("ERR_Code") = 3
  Response.Redirect("login.asp")
End If

  sUsername = Trim(Request.Form("cms_anvnamn"))
  sPassword = Request.Form("cms_passwd")

  If Len(sUsername) < 1 Or Len(sUsername) > 50 Then
    Session.Value("ERR_Code") = 1
    Response.Redirect("login.asp")
  End If
  
  Con_Open()
    
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open "SELECT * FROM " & LOGIN_TABLE & " WHERE " & LOGIN_USR & " = '" & ITrim(sUsername) & "'", con, 1, 3
    
      If rs.EOF Then
        Session.Value("ERR_Code") = 1
        Response.Redirect("login.asp")
      End If
      
      If rs.RecordCount > 1 Then
        Session.Value("ERR_Code") = 1
        Response.Redirect("login.asp")
      End If
      
      'If Not rs("aS_System") And Not rs("aS_Admin") And Not rs("aS_Redaktion") Then
      '  Session.Value("ERR_Code") = 1
      '  Response.Redirect("login.asp")
      'End If
      
      If Not rs("aS_CMS") Then
        Session.Value("ERR_Code") = 5
        Response.Redirect("login.asp")
      End If
    
      sDBSalt1  = rs("aSalt1")
      sDBSalt2  = rs("aSalt2")
      sHash     = config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPassword & "" & config_Hash_Salt_2 & "" & sDBSalt2
      sHash     = MD5(sHash)
      
      If Not rs(LOGIN_PWS) = sHash Then
        Session.Value("ERR_Code") = 1
        Response.Redirect("login.asp")
      End If
      
      ' #### OK GODKÄND, SÄTT VARIABLER ####
      
      Session.Value("CMS_LOGIN")  = True
      Session.Value("CMS_ID")     = rs("aID")
      Session.Value("CMS_ANAMN")  = rs("aAnvNamn")
      Session.Value("CMS_NAMN")   = rs("aNamn")
      
      Session.Value("CMS_RATTER") =  rs("aS_CMSRatter")
      
      Session.Value("CMS_sSYS")   =  rs("aS_System")
      Session.Value("CMS_sADM")   =  rs("aS_Admin")
      Session.Value("CMS_sRED")   =  rs("aS_Redaktion")
      
      
      
    rs.Close
    Set rs = Nothing
  
  Con_Close()

Response.Redirect("/")
%>