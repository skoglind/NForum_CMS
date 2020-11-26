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
    sVD        = mForm("vVD", "ABC", 255)
    sHemsida   = mForm("vHemsida", "ABC", 255)
    sHemland   = mForm("vHemland", "ABC", 255)
    sNyckelord = mForm("vNyckelord", "ABC", 500)
    sTextM     = mForm("vText", "ABC", 0)
    
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Foretag WHERE fID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sNamn) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett något namn för företaget.</p>"
        
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
          eBild = GetImgIDByExclusive()
          If eBild > 0 Then
            rsDB("fLogga") = eBild
            Con.ExeCute("UPDATE cms_Bild SET bSparad = 1 WHERE bID = " & CLng(eBild))
          End If
        End If
        
        rsDB("fNamn")      = sNamn
        rsDB("fVD")        = sVD
        
        If Trim(sHemsida)  = "http://" Then sHemsida = ""
        rsDB("fHemsida")   = sHemsida
        rsDB("fHemland")   = sHemland
        rsDB("fNyckelord") = sNyckelord
        rsDB("fTextM")     = sTextM
        
        rsDB("fDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("fID")
        If bIsNew Then AddLogg "FÖRETAG","SKAPA",lID
        
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
  Case "glue" ' Slå ihop och radera
    myID  = mForm("fMy_ID", "123", 0)
    addID = mForm("fAdd_ID", "123", 0)
  
    If GetAcc("CMS44") And addID > 0 And myID > 0 Then
      Con_Open

        Con.ExeCute("UPDATE cms_Speltitlar SET tUtgivare = " & CLng(addID) & " WHERE tUtgivare = " & CLng(myID))
        Con.ExeCute("UPDATE cms_Spel SET sUtvecklare = " & CLng(addID) & " WHERE sUtvecklare = " & CLng(myID))
        Con.ExeCute("DELETE FROM cms_Bild WHERE bID IN (SELECT fLogga FROM cms_Foretag WHERE fID = " & CLng(myID) & ")")
        
        fLogga = Con.ExeCute("SELECT fLogga FROM cms_Foretag WHERE fID = " & CLng(myID))(0)
        If Not IsNumeric(fLogga) Or fLogga = Empty Then fLogga = 0
        Con.ExeCute("DELETE FROM cms_Foretag WHERE fID = " & CLng(myID))
        ImgRemove fLogga
        
      Con_CLose
      Response.Write("<script type='text/javascript'>parent.gluefinished('" & "_edit.asp?e=" & addID & "&" & sRebuild & "');</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.gluefailed();</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
  Case "del" ' Radera
    If GetAcc("CMS44") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM cms_Foretag WHERE fID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              ImgRemove rsDB("fLogga")
              Call AddLogg("FÖRETAG","RADERA [TOTAL]",rsDB("fID"))
              rsDB.Delete
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade företag som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Databas/Foretag/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>