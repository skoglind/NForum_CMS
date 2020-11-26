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
    lSpelID     = mForm("vSpelID", "123", 0)
    lBetyg      = mForm("vBetyg", "123", 0)
    sTitel      = mForm("vTitel", "ABC", 255)
    sTextM      = mForm("vText", "ABC", 0)
    sNotes      = mForm("vNotes_cp", "ABC", 0)
    sShort      = mForm("vShort", "ABC", 100)
    sNyckelord  = mForm("vNyckelord", "ABC", 500)
    bAnvRec     = mForm("vAnvRec", "CHK", 0)
    
    Con_Open
      If NOT GetAcc("CMS111") Then sbFilter = " AND NOT rStatus = 0"
    
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Recensioner WHERE rID = " & CLng(lID) & sbFilter
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If lKategori = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du valt en otillåten kategori.</p>"
        If Len(sTitel) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon titel för recensionen.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
          If Not GetAcc("CMS111") And CLng(rsDB("rSkapadAv")) <> CLng(cCMS_ID) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du saknar behörighet att ändra denna recension.</p>"
          If Not GetAcc("CMS11") And (CLng(rsDB("rStatus")) = CLng(4) Or CLng(rsDB("rStatus")) = CLng(2)) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du saknar behörighet att ändra denna recension.</p>"
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
          rsDB("rDatumSkapad")  = Now
          rsDB("rSkapadAv")     = cCMS_ID
          rsDB("rStatus")       = 1
        Else
          lNySkapare  = mForm("vNySkapare_cp", "123", 0)
          If lNySkapare > 0 Then If lNySkapare <> CLng(rsDB("rSkapadAv")) And GetAcc("CMS111") Then rsDB("rSkapadAv") = lNySkapare
          
          ' #### ÄNDRA STATUS ####
          lStatus     = mForm("vStatus_cp", "123", 0)
          lPublTyp    = mForm("vPublTyp_cp", "123", 0)
          lPublDatum  = mForm("vPublDatum_cp", "DAT", 0)
          
          nStatus = rsDB("rStatus")
          
          If lStatus <> 0 And GetAcc("CMS11") Then
            Select Case lStatus
              Case 1  ' Under bearbetning
                nStatus = 1
                Call AddLogg("RECENSION","PUBLICERING [ÅTER BEARBETNING]",lID)
              Case 2  ' Inväntar publicering
                nStatus = 2
                Call AddLogg("RECENSION","PUBLICERING [INVÄNTA]",lID)
              Case 3  ' Publicering nekad
                nStatus = 3
                Call AddLogg("RECENSION","PUBLICERING [NEKAD]",lID)
              Case 4  ' Publicerad
                nStatus = 4
                
                ' #### PUBLICERA, VÄLJ DATUM ####
                Select Case lPublTyp
                  Case 1  ' DIREKT
                    rsDB("rDatumPublicerad") = Now
                    Call AddLogg("RECENSION","PUBLICERING [GODKÄNN DIREKT]",lID)
                  Case 2  ' IMORGON
                    rsDB("rDatumPublicerad") = CDate(DateAdd("d", 1, Date) & " " & PUBL_TID)
                    Call AddLogg("RECENSION","PUBLICERING [GODKÄNN IMORGON]",lID)
                  Case 3  ' NÄSTA VARDAG
                    Do Until IsVardag
                      dPublicerad = DateAdd("d", 1, Date)
                      If weekday(dPublicerad) <> 6 And weekday(dPublicerad) <> 7 Then IsVardag = True
                    Loop
                    rsDB("rDatumPublicerad") = CDate(dPublicerad & " " & PUBL_TID)
                    Call AddLogg("RECENSION","PUBLICERING [GODKÄNN VARDAG]",lID)
                  Case 4  ' VALT DATUM
                    If Not lPublDatum > #2025-01-01# Then 
                      rsDB("rDatumPublicerad") = CDate(lPublDatum & " " & PUBL_TID)
                      Call AddLogg("RECENSION","PUBLICERING [GODKÄNN DATUM]",lID)
                    Else
                      bDontChange = True
                    End If
                  Case Else
                    bDontChange = True
                End Select
                 
                If nStatus = 4 And Not bDontChange Then rsDB("rPubliceradAv") = cCMS_ID
                ' ###############################
                
              Case Else
                bDontChange = True
            End Select
          End If
          
          If Not bDontChange Then rsDB("rStatus") = nStatus
          ' ######################
        End If
        
        rsDB("rKategori")   = lKategori
        rsDB("rTitel")      = sTitel
        rsDB("rText")       = sTextM
        rsDB("rNotes")      = sNotes
        rsDB("rShort")      = sShort
        rsDB("rNyckelord")  = sNyckelord
        rsDB("rBetyg")      = lBetyg
        rsDB("rAnvandarRec")= bAnvRec
        rsDB("rSpelID")     = CLng(lSpelID)
        
        rsDB("rDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("rID")
        If bIsNew Then AddLogg "RECENSION","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
      
      If bIsNew Then Con.ExeCute("UPDATE cms_Bind_Recension_Img SET brRecension = " & CLng(lID) & ", brSaved = 1 WHERE brSaved = 0 And brUser = " & CLng(cCMS_ID))
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
            If Not GetAcc("CMS111") Then sbFilter = " AND rSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Recensioner WHERE rID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("rStatus") = 1 Then
                rsDB("rStatus") = 2
                rsDB.Update
                Call AddLogg("RECENSION","PUBLICERING [INVÄNTA]",rsDB("rID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Statusändring slutförd</h2><p>De markerade recensionerna som du hade behörighet att ändra status på har nu ändrats till ""Inväntar publicering"".</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Recensioner/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "unawait" ' Åter under bearbetning
    If GetAcc("CMS1") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            If Not GetAcc("CMS111") Then sbFilter = " AND rSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Recensioner WHERE rID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("rStatus") = 2 Then
                rsDB("rStatus") = 1
                rsDB.Update
                Call AddLogg("RECENSION","PUBLICERING [ÅTER BEARBETNING]",rsDB("rID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Statusändring slutförd</h2><p>De markerade recensionerna som du hade behörighet att ändra status på har nu ändrats till ""Under bearbetning"".</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Recensioner/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "del" ' Radera
    Con_Open
      allID = Split(GetFormRequest("chk_id", "YES"), ",")
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      
        For Each oID IN allID
          If Not GetAcc("CMS111") Then sbFilter = " AND rSkapadAv = " & cCMS_ID
          SQL = "SELECT * FROM cms_Recensioner WHERE rID = " & CLng(oID) & sbFilter
          rsDB.Open SQL, Con, 1, 3
        
          If Not rsDB.Eof Then
            If rsDB("rStatus") = 0 And GetAcc("CMS111") Then
              Call AddLogg("RECENSION","RADERA [TOTAL]",rsDB("rID"))
              
              Set rsBilder = Server.CreateObject("ADODB.RecordSet")
              SQL = "SELECT * FROM cms_Bind_Recension_Img WHERE brRecension = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("brBild")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close
              Set rsBilder = Nothing
              
              Con.ExeCute("DELETE FROM cms_Bind_Recension_Img WHERE brRecension = " & CLng(oID))
              rsDB.Delete
            ElseIf rsDB("rStatus") = 4 Or rsDB("rStatus") = 2 Then
              If GetAcc("CMS11") Then
                rsDB("rStatus") = 0
                rsDB.Update
                Call AddLogg("RECENSION","RADERA [ENKEL]",rsDB("rID"))
              End If
            Else
              rsDB("rStatus") = 0
              rsDB.Update
              Call AddLogg("RECENSION","RADERA [ENKEL]",rsDB("rID"))
            End If
          End If
          
          rsDB.Close
        Next
      
      Set rsDB = Nothing
    Con_Close
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade recensionerna som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Recensioner/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>