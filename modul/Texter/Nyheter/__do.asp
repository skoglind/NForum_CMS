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
    sKalla      = mForm("vKalla", "ABC", 500)
    sNyckelord  = mForm("vNyckelord", "ABC", 500)
    
    Con_Open
      If NOT GetAcc("CMS111") Then sbFilter = " AND NOT nStatus = 0"
    
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Nyheter WHERE nID = " & CLng(lID) & sbFilter
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If lKategori = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen d� du valt en otill�ten kategori.</p>"
        If Len(sTitel) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen d� du inte har angett n�gon titel f�r nyheten.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
          If Not GetAcc("CMS111") And CLng(rsDB("nSkapadAv")) <> CLng(cCMS_ID) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen d� du saknar beh�righet att �ndra denna nyhet.</p>"
          If Not GetAcc("CMS11") And (CLng(rsDB("nStatus")) = CLng(4) Or CLng(rsDB("nStatus")) = CLng(2)) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen d� du saknar beh�righet att �ndra denna nyhet.</p>"
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
          rsDB("nDatumSkapad")  = Now
          rsDB("nDatumPublicerad") = "2050-01-01 00:00:00"
          rsDB("nSkapadAv")     = cCMS_ID
          rsDB("nStatus")       = 1
        Else
          lNySkapare  = mForm("vNySkapare_cp", "123", 0)
          If lNySkapare > 0 Then If lNySkapare <> CLng(rsDB("nSkapadAv")) And GetAcc("CMS111") Then rsDB("nSkapadAv") = lNySkapare
          
          ' #### �NDRA STATUS ####
          lStatus     = mForm("vStatus_cp", "123", 0)
          lPublTyp    = mForm("vPublTyp_cp", "123", 0)
          lPublDatum  = mForm("vPublDatum_cp", "DAT", 0)
          
          nStatus = rsDB("nStatus")
          
          If lStatus <> 0 And GetAcc("CMS11") Then
            Select Case lStatus
              Case 1  ' Under bearbetning
                nStatus = 1
                Call AddLogg("NYHET","PUBLICERING [�TER BEARBETNING]",lID)
              Case 2  ' Inv�ntar publicering
                nStatus = 2
                Call AddLogg("NYHET","PUBLICERING [INV�NTA]",lID)
              Case 3  ' Publicering nekad
                nStatus = 3
                Call AddLogg("NYHET","PUBLICERING [NEKAD]",lID)
              Case 4  ' Publicerad
                nStatus = 4
                
                ' #### PUBLICERA, V�LJ DATUM ####
                Select Case lPublTyp
                  Case 1  ' DIREKT
                    rsDB("nDatumPublicerad") = Now
                    Call AddLogg("NYHET","PUBLICERING [GODK�NN DIREKT]",lID)
                  Case 2  ' IMORGON
                    rsDB("nDatumPublicerad") = CDate(DateAdd("d", 1, Date) & " " & PUBL_TID)
                    Call AddLogg("NYHET","PUBLICERING [GODK�NN IMORGON]",lID)
                  Case 3  ' N�STA VARDAG
                    Do Until IsVardag
                      dPublicerad = DateAdd("d", 1, Date)
                      If weekday(dPublicerad) <> 6 And weekday(dPublicerad) <> 7 Then IsVardag = True
                    Loop
                    rsDB("nDatumPublicerad") = CDate(dPublicerad & " " & PUBL_TID)
                    Call AddLogg("NYHET","PUBLICERING [GODK�NN VARDAG]",lID)
                  Case 4  ' VALT DATUM
                    If Not lPublDatum > #2025-01-01# Then 
                      rsDB("nDatumPublicerad") = CDate(lPublDatum & " " & PUBL_TID)
                      Call AddLogg("NYHET","PUBLICERING [GODK�NN DATUM]",lID)
                    Else
                      bDontChange = True
                    End If
                  Case Else
                    bDontChange = True
                End Select
                 
                If nStatus = 4 And Not bDontChange Then rsDB("nPubliceradAv") = cCMS_ID
                ' ###############################
                
              Case Else
                bDontChange = True
            End Select
          End If
          
          If Not bDontChange Then rsDB("nStatus") = nStatus
          ' ######################
        End If
        
        rsDB("nKategori")   = lKategori
        rsDB("nTitel")      = sTitel
        rsDB("nText")       = sTextM
        rsDB("nNotes")      = sNotes
        rsDB("nKalla")      = sKalla
        rsDB("nNyckelord")  = sNyckelord
        
        rsDB("nDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"

        rsDB.Update
        
        lID     = rsDB("nID")
        If bIsNew Then AddLogg "NYHET","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
      
      If bIsNew Then Con.ExeCute("UPDATE cms_Bind_Nyheter_Img SET bnNyhet = " & CLng(lID) & ", bnSaved = 1 WHERE bnSaved = 0 And bnUser = " & CLng(cCMS_ID))
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
  Case "await" ' Inv�nta publicering
    If GetAcc("CMS1") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            If Not GetAcc("CMS111") Then sbFilter = " AND nSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Nyheter WHERE nID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("nStatus") = 1 Then
                rsDB("nStatus") = 2
                rsDB.Update
                Call AddLogg("NYHET","PUBLICERING [INV�NTA]",rsDB("nID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Status�ndring slutf�rd</h2><p>De markerade nyheter som du hade beh�righet att �ndra status p� har nu �ndrats till ""Inv�ntar publicering"".</p><p>Klicka p� ""forts�tt"" f�r att g� vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Nyheter/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "unawait" ' �ter under bearbetning
    If GetAcc("CMS1") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            If Not GetAcc("CMS111") Then sbFilter = " AND nSkapadAv = " & cCMS_ID
            SQL = "SELECT * FROM cms_Nyheter WHERE nID = " & CLng(oID) & sbFilter
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              If rsDB("nStatus") = 2 Then
                rsDB("nStatus") = 1
                rsDB.Update
                Call AddLogg("NYHET","PUBLICERING [�TER BEARBETNING]",rsDB("nID"))
              End If
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Status�ndring slutf�rd</h2><p>De markerade nyheter som du hade beh�righet att �ndra status p� har nu �ndrats till ""Under bearbetning"".</p><p>Klicka p� ""forts�tt"" f�r att g� vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Nyheter/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "del" ' Radera
    Con_Open
      allID = Split(GetFormRequest("chk_id", "YES"), ",")
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      
        For Each oID IN allID
          If Not GetAcc("CMS111") Then sbFilter = " AND nSkapadAv = " & cCMS_ID
          SQL = "SELECT * FROM cms_Nyheter WHERE nID = " & CLng(oID) & sbFilter
          rsDB.Open SQL, Con, 1, 3
        
          If Not rsDB.Eof Then
            If rsDB("nStatus") = 0 And GetAcc("CMS111") Then
              Call AddLogg("NYHET","RADERA [TOTAL]",rsDB("nID"))
              
              Set rsBilder = Server.CreateObject("ADODB.RecordSet")
              SQL = "SELECT * FROM cms_Bind_Nyheter_Img WHERE bnNyhet = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("bnBild")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close
              Set rsBilder = Nothing
              
              Con.ExeCute("DELETE FROM cms_Bind_Nyheter_Img WHERE bnNyhet = " & CLng(oID))
              
              rsDB.Delete
            ElseIf rsDB("nStatus") = 4 Or rsDB("nStatus") = 2 Then
              If GetAcc("CMS11") Then
                rsDB("nStatus") = 0
                rsDB.Update
                Call AddLogg("NYHET","RADERA [ENKEL]",rsDB("nID"))
              End If
            Else
              rsDB("nStatus") = 0
              rsDB.Update
              Call AddLogg("NYHET","RADERA [ENKEL]",rsDB("nID"))
            End If
          End If
          
          rsDB.Close
        Next
      
      Set rsDB = Nothing
    Con_Close
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutf�rd</h2><p>De markerade nyheter som du hade beh�righet att radera �r nu borta.</p><p>Klicka p� ""forts�tt"" f�r att g� vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Nyheter/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>