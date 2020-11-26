<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS4") Then Response.Redirect("/")
%>

<%
sAction       = Trim(LCase(Request.QueryString("a")))
sExtraAction  = Trim(LCase(Request.QueryString("ea")))

' #### REMEMBER ####
sFilter = noFnutt(Request.Form("f"))
lPaSida = noFnutt(Request.Form("s"))
Call GetAlfa(Request.Form("alfa"))

sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa
' ##################

Select Case sAction
  Case "save" ' Spara
    lID        = mForm("vID", "123", 0)
    sNamn      = mForm("vNamn", "ABC", 255)
    sNyckelord = mForm("vNyckelord", "ABC", 500)
    sTextM     = mForm("vText", "ABC", 0)
    sSerien    = mForm("vSerien", "CHK", 0)
    
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Spelserier WHERE ssID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sNamn) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett något namn för spelgruppen.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
        Else
          bIsNew = True
        End If
        
        If bErr Then
          Response.Write("<script type='text/javascript'>parent.savefailed('" & nMessage & "');</script>")
          Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
          Response.End
        End If
      ' ######################
      
        If rsDB.EOF Then
          rsDB.AddNew
        End If
        
        rsDB("ssNamn")      = sNamn
        rsDB("ssNyckelord") = sNyckelord
        rsDB("ssTextM")     = sTextM
        If sSerien Then rsDB("ssSerien") = True Else rsDB("ssSerien") = False
        
        rsDB("ssDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("ssID")
        If bIsNew Then AddLogg "SPELGRUPP","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
    Con_Close
    
    Select Case sExtraAction
      Case "continue"
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",true,'" & "_edit.asp?" & sRebuild & "');</script>")
      Case "return"
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",true,'" & "_show.asp?" & sRebuild & "');</script>")
      Case Else
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",false,'');</script>")
    End Select
    
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
  Case "del" ' Radera
    If GetAcc("CMS44") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM cms_Spelserier WHERE ssID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              Call AddLogg("SPELGRUPP","RADERA [TOTAL]",rsDB("ssID"))
              Con.ExeCute("DELETE FROM cms_SpelserieBind_Spel WHERE bsSpelSerie = " & CLng(oID))
              rsDB.Delete
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade spelgrupper som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Databas/Spelserier/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>