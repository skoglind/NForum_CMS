<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS1") Then Response.Redirect("/")
%>

<%
sAction       = Trim(LCase(Request.QueryString("a")))
sExtraAction  = Trim(LCase(Request.QueryString("ea")))

' #### REMEMBER ####
  sFilter = noFnutt(Request.Form("f"))
  lPaSida = noFnutt(Request.Form("s"))
  sRebuild = "f=" & sFilter & "&s=" & lPaSida
' ##################

Select Case sAction
  Case "save" ' Spara
    lID         = mForm("vID", "123", 0)
    lKategori   = mForm("vKategori", "123", 0)
    sTitel      = mForm("vTitel", "ABC", 255)
    sTextM      = mForm("vText", "ABC", 0)
    sNotes      = mForm("vNotes_cp", "ABC", 0)
    sShort      = mForm("vShort", "ABC", 100)
    sNyckelord  = mForm("vNyckelord", "ABC", 500)
    bAnvArt     = mForm("vAnvArt", "CHK", 0)
    
    Con_Open
      If NOT GetAcc("CMS111") Then sbFilter = " AND NOT aaStatus = 0"
    
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Artiklar WHERE aaID = " & CLng(lID) & sbFilter
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If lKategori = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du valt en otillåten kategori.</p>"
        If Len(sTitel) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon titel för artikeln.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
          If Not GetAcc("CMS111") And CLng(rsDB("aaSkapadAv")) <> CLng(cCMS_ID) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du saknar behörighet att ändra denna artikel.</p>"
          If Not GetAcc("CMS11") And (CLng(rsDB("aaStatus")) = CLng(4) Or CLng(rsDB("aaStatus")) = CLng(2)) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du saknar behörighet att ändra denna artikel.</p>"
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
          rsDB("aaDatumSkapad")  = Now
          rsDB("aaSkapadAv")     = cCMS_ID
          rsDB("aaStatus")       = 1
        Else
          lNySkapare  = mForm("vNySkapare_cp", "123", 0)
          If lNySkapare > 0 Then If lNySkapare <> CLng(rsDB("aaSkapadAv")) And GetAcc("CMS111") Then rsDB("aaSkapadAv") = lNySkapare
          
          ' #### ÄNDRA STATUS ####
          lStatus     = mForm("vStatus_cp", "123", 0)
          lPublTyp    = mForm("vPublTyp_cp", "123", 0)
          lPublDatum  = mForm("vPublDatum_cp", "DAT", 0)
          
          nStatus = rsDB("aaStatus")
          
          If lStatus <> 0 And GetAcc("CMS11") Then
            Select Case lStatus
              Case 1  ' Under bearbetning
                nStatus = 1
                Call AddLogg("ARTIKEL","PUBLICERING [ÅTER BEARBETNING]",lID)
              Case 2  ' Inväntar publicering
                nStatus = 2
                Call AddLogg("ARTIKEL","PUBLICERING [INVÄNTA]",lID)
              Case 3  ' Publicering nekad
                nStatus = 3
                Call AddLogg("ARTIKEL","PUBLICERING [NEKAD]",lID)
              Case 4  ' Publicerad
                nStatus = 4
                
                ' #### PUBLICERA, VÄLJ DATUM ####
                Select Case lPublTyp
                  Case 1  ' DIREKT
                    rsDB("aaDatumPublicerad") = Now
                    Call AddLogg("ARTIKEL","PUBLICERING [GODKÄNN DIREKT]",lID)
                  Case 2  ' IMORGON
                    rsDB("aaDatumPublicerad") = CDate(DateAdd("d", 1, Date) & " " & PUBL_TID)
                    Call AddLogg("ARTIKEL","PUBLICERING [GODKÄNN IMORGON]",lID)
                  Case 3  ' NÄSTA VARDAG
                    Do Until IsVardag
                      dPublicerad = DateAdd("d", 1, Date)
                      If weekday(dPublicerad) <> 6 And weekday(dPublicerad) <> 7 Then IsVardag = True
                    Loop
                    rsDB("aaDatumPublicerad") = CDate(dPublicerad & " " & PUBL_TID)
                    Call AddLogg("ARTIKEL","PUBLICERING [GODKÄNN VARDAG]",lID)
                  Case 4  ' VALT DATUM
                    If Not lPublDatum > #2025-01-01# Then 
                      rsDB("aaDatumPublicerad") = CDate(lPublDatum & " " & PUBL_TID)
                      Call AddLogg("ARTIKEL","PUBLICERING [GODKÄNN DATUM]",lID)
                    Else
                      bDontChange = True
                    End If
                  Case Else
                    bDontChange = True
                End Select
                 
                If nStatus = 4 And Not bDontChange Then rsDB("aaPubliceradAv") = cCMS_ID
                ' ###############################
                
              Case Else
                bDontChange = True
            End Select
          End If
          
          If Not bDontChange Then rsDB("aaStatus") = nStatus
          ' ######################
        End If
        
        rsDB("aaKategori")   = lKategori
        rsDB("aaTitel")      = sTitel
        rsDB("aaText")       = sTextM
        rsDB("aaNotes")      = sNotes
        rsDB("aaShort")      = sShort
        rsDB("aaNyckelord")  = sNyckelord
        rsDB("aaAnvandarArt")= bAnvArt
        
        rsDB("aaDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("aaID")
        If bIsNew Then AddLogg "ARTIKEL","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
      
      If bIsNew Then Con.ExeCute("UPDATE cms_Bind_Artikel_Img SET baArtikel = " & CLng(lID) & ", baSaved = 1 WHERE baSaved = 0 And baUser = " & CLng(cCMS_ID))
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
  Case "await" ' Invänta publicering
    If GetAcc("CMS1") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            If Not GetAcc("CMS111") Then sbFilter = " AND aaSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Artiklar WHERE aaID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("aaStatus") = 1 Then
                rsDB("aaStatus") = 2
                rsDB.Update
                Call AddLogg("ARTIKEL","PUBLICERING [INVÄNTA]",rsDB("aaID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Statusändring slutförd</h2><p>De markerade artiklarna som du hade behörighet att ändra status på har nu ändrats till ""Inväntar publicering"".</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Artiklar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "unawait" ' Åter under bearbetning
    If GetAcc("CMS1") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            If Not GetAcc("CMS111") Then sbFilter = " AND aaSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Artiklar WHERE aaID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("aaStatus") = 2 Then
                rsDB("aaStatus") = 1
                rsDB.Update
                Call AddLogg("ARTIKEL","PUBLICERING [ÅTER BEARBETNING]",rsDB("aaID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Statusändring slutförd</h2><p>De markerade artiklarna som du hade behörighet att ändra status på har nu ändrats till ""Under bearbetning"".</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Artiklar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "del" ' Radera
    Con_Open
      allID = Split(GetFormRequest("chk_id", "YES"), ",")
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      
        For Each oID IN allID
          If Not GetAcc("CMS111") Then sbFilter = " AND aaSkapadAv = " & cCMS_ID
          SQL = "SELECT * FROM cms_Artiklar WHERE aaID = " & CLng(oID) & sbFilter
          rsDB.Open SQL, Con, 1, 3
        
          If Not rsDB.Eof Then
            If rsDB("aaStatus") = 0 And GetAcc("CMS111") Then
              Call AddLogg("ARTIKEL","RADERA [TOTAL]",rsDB("aaID"))
              
              Set rsBilder = Server.CreateObject("ADODB.RecordSet")
              SQL = "SELECT * FROM cms_Bind_Artikel_Img WHERE baArtikel = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("baBild")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close
              Set rsBilder = Nothing
              
              Con.ExeCute("DELETE FROM cms_Bind_Artikel_Img WHERE baArtikel = " & CLng(oID))
              rsDB.Delete
            ElseIf rsDB("aaStatus") = 4 Or rsDB("aaStatus") = 2 Then
              If GetAcc("CMS11") Then
                rsDB("aaStatus") = 0
                rsDB.Update
                Call AddLogg("ARTIKEL","RADERA [ENKEL]",rsDB("aaID"))
              End If
            Else
              rsDB("aaStatus") = 0
              rsDB.Update
              Call AddLogg("ARTIKEL","RADERA [ENKEL]",rsDB("aaID"))
            End If
          End If
          
          rsDB.Close
        Next
      
      Set rsDB = Nothing
    Con_Close
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade artiklarna som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Artiklar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>