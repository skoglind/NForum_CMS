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
sQ = Trim(Left(MakeLegal_Large(Request.Form("q")), 255))
sFilter = noFnutt(Request.Form("f"))
lPaSida = noFnutt(Request.Form("s"))
Call GetAlfa(Request.Form("alfa"))

sRebuild = "f=" & sFilter & "&s=" & lPaSida & "&alfa=" & sSendAlfa & "&q=" & sQ
' ##################

Select Case sAction
  Case "save" ' Spara
    lID         = mForm("vID", "123", 0)
    sKonsol     = mForm("vKonsol", "123", 0)
    sForetag    = mForm("vUtvecklare", "123", 0)
    sNyckelord  = mForm("vNyckelord", "ABC", 500)
    sTextM      = mForm("vText", "ABC", 0)
    
    sSynlig     = mForm("vSynlig_cp", "CHK", 0)
    sESRB       = mForm("vESRB_cp", "123", 0)
    sPEGI       = mForm("vPEGI_cp", "123", 0)
    sSpelare    = mForm("vAntalSpelare_cp", "123", 0)
    sSinglePlay = mForm("vSinglePlay_cp", "CHK", 0)
    sMultiPlay  = mForm("vMultiPlay_cp", "CHK", 0)
    sOnline     = mForm("vOnline_cp", "CHK", 0)
    sLicense    = mForm("vLicense_cp", "CHK", 0)
    
    nID         = FormIDToArray("vRegion_")
    nUseTitel   = mForm("vStandardTitel", "ABC", 50)
    
    ReDim nRegion(UBound(nID))
    ReDim nSortNo(UBound(nID))
    ReDim nTitel(UBound(nID))
    ReDim nExtra(UBound(nID))
    ReDim nRelease(UBound(nID))
    ReDim nRegionskod(UBound(nID))
    ReDim nUtgivareID(UBound(nID))
    ReDim nAction(UBound(nID))
    
    For zx = 0 To UBound(nID)
      nID(zx) = Trim(nID(zx))
    
      nRegion(zx)     = mForm("vRegion_" & nID(zx), "123", 0)
      nSortNo(zx)     = mForm("vSortNo_" & nID(zx), "123", 0)
      nTitel(zx)      = mForm("vTitel_" & nID(zx), "ABC", 255)
      nExtra(zx)      = mForm("vExtra_" & nID(zx), "ABC", 255)
      nRelease(zx)    = mForm("vRelease_" & nID(zx), "ABC", 40)
      nRegionskod(zx) = mForm("vRegionskod_" & nID(zx), "ABC", 50)
      nUtgivareID(zx) = mForm("vUtgivareID_" & nID(zx), "123", 0)
      nAction(zx)     = False
    Next
     
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Spel WHERE sID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If sKonsol = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du valt en otillåten konsol.</p>"
        If UBound(nID) < 0 Then bErr = True : nMessage = "<p>Om du ser detta meddelande beror det på att något gick fel med att ladda in titlarna, sparningen är stoppad för säkerhets skull. Allt du behöver göra är att spara på nytt.</p>"
        
        For zx = 0 To UBound(nID)
          If Trim(nTitel(zx)) = Empty Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte fyllt i en titel (observera att alla regioner måste ha sin titel ifylld).</p>"
        Next
        
        For zx = 0 To UBound(nID)
          If Not IsNumeric(nRegion(zx)) Or nRegion(zx) = Empty Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte fyllt i en giltig region.</p>"
        Next
        
        For zx = 0 To UBound(nID)
          If Not IsNumeric(nUtgivareID(zx)) Or nUtgivareID(zx) = Empty Then 
            nUtgivareID(zx) = 0
          Else
            nUtgivareID(zx) = CLng(nUtgivareID(zx))
          End If
        Next
       
        If Trim(nUseTitel & " ") = Empty Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte valt en standardtitel.</p>"
        
        For zx = 0 To UBound(nID)
          nRelease(zx) = LCase(Trim(nRelease(zx)))
        
          If IsDate(nRelease(zx)) Then
            nRelease(zx) = Cdate(nRelease(zx)) 
          ElseIf Len(nRelease(zx)) = 4 And IsNumeric(nRelease(zx)) Then
            nRelease(zx) = nRelease(zx)
          ElseIf nRelease(zx) = "n/r" Then
            nRelease(zx) = "N/R"
          ElseIf Len(nRelease(zx)) = 0 Or nRelease(zx) = "n/a" Then
            nRelease(zx) = "N/A"
          Else
            bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte fyllt i ett giltigt datum.</p>"
          End If
        Next
        
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
        
        rsDB("sUtvecklare")   = CLng(sForetag)
        rsDB("sKonsol")       = CLng(sKonsol)
        rsDB("sNyckelord")    = sNyckelord
        rsDB("sTextM")        = sTextM
        
        rsDB("sSynlig")       = sSynlig
        rsDB("sESRB")         = CLng(sESRB)
        rsDB("sPEGI")         = CLng(sPEGI)
        rsDB("sAntalSpelare") = CLng(sSpelare)
        rsDB("sSingleplayer") = sSinglePlay
        rsDB("sMultiplayer")  = sMultiPlay
        rsDB("sOlicensierad") = sLicense
        rsDB("sOnline")       = sOnline
        
        rsDB("sDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
        
        lID = rsDB("sID")
        If bIsNew Then AddLogg "SPEL","SKAPA",lID
        
        ' #### TITLAR ####
          Set rsT = Server.CreateObject("ADODB.RecordSet")
          SQL = "SELECT * FROM cms_Speltitlar WHERE tSpelID = " & CLng(lID) & " ORDER BY tID ASC"
          rsT.Open SQL, Con, 1, 3
          
          lFindID = -1
          Do Until rsT.EOF
            lFindID = IsInArray(nID, rsT("tID"))
            If lFindID < 0 Then lFindID = IsInArray(nID, rsT("tSparadNyckel"))
            
            If lFindID < 0 Then
              ImgRemove rsT("tBoxart_BoxFram")
              ImgRemove rsT("tBoxart_BoxBak")
              ImgRemove rsT("tBoxart_Manual")
              ImgRemove rsT("tBoxart_Kassett")
            
              rsT.Delete
            Else
              nAction(lFindID)    = True
              rsT("tUtgivare")    = CLng(nUtgivareID(lFindID))
              rsT("tTitel")       = Left(Trim(nTitel(lFindID) & " "),255)
              rsT("tExtra")       = Left(Trim(nExtra(lFindID) & " "),255)
              rsT("tRelease")     = Left(nRelease(lFindID),25)
              rsT("tRegion")      = CLng(nRegion(lFindID))
              rsT("tSortNo")      = CLng(nSortNo(lFindID))
              rsT("tRegionsKod")  = Left(Trim(nRegionskod(lFindID) & " "),50)
              
              rsT.Update
              
              If MakeLegal(nUseTitel) = MakeLegal(nID(lFindID)) Then nUseTitel = rsT("tID")
            End If
          
            rsT.MoveNext
          Loop
          
          rsT.Close
          
          For zx = 0 To UBound(nID)
            If Not nAction(zx) Then
              SQL = "SELECT * FROM cms_Speltitlar WHERE tSparadNyckel = '" & MakeLegal(nID(zx)) & "' AND tSparadAv = " & CLng(cCMS_ID)
              rsT.Open SQL, Con
              
              If rsT.EOF Then
                rsT.AddNew
              End If
              
              nAction(zx)           = True
              rsT("tSpelID")        = lID
              rsT("tUtgivare")      = CLng(nUtgivareID(zx))
              rsT("tTitel")         = Left(Trim(nTitel(zx)),255)
              rsT("tExtra")         = Left(Trim(nExtra(zx)),255)
              rsT("tRelease")       = Left(nRelease(zx),25)
              rsT("tRegion")        = CLng(nRegion(zx))
              rsT("tSortNo")        = CLng(nSortNo(zx))
              rsT("tRegionsKod")    = Left(Trim(nRegionskod(zx)),50)
              rsT("tSparadNyckel")  = Left(Trim(MakeLegal(nID(zx))),50)
              rsT("tSparadAv")      = CLng(cCMS_ID)
              
              rsT.Update
              
              If MakeLegal(nUseTitel) = MakeLegal(nID(zx)) Then nUseTitel = rsT("tID")
              
              rsT.Close
            End If
          Next
          
          Set rsT = Nothing
          
          Con.ExeCute("UPDATE cms_Spel SET sStandard_Titel = " & CLng(nUseTitel) & " WHERE sID = " & CLng(lID))
        ' ################
        
        ' #### GENRE ####
          For Each lGenre In Request.Form("genre")
            sGenre = sGenre & lGenre & ","
          Next
          If Len(sGenre) > 0 Then sGenre = Left(sGenre, Len(sGenre)-1)
          If sGenre = "" Then sGenre = "0"
          aGenre = Split(sGenre, ",")
        ' ###############
          
        ' #### GRUPP ####
          For Each lGrupp In Request.Form("grupp")
            sGrupp = sGrupp & lGrupp & ","
          Next
          If Len(sGrupp) > 0 Then sGrupp = Left(sGrupp, Len(sGrupp)-1)
          If sGrupp = "" Then sGrupp = "0"
          aGrupp = Split(sGrupp, ",")
        ' ###############
          
        ' #### LISTOR ####
          Con.ExeCute("DELETE FROM cms_Bind_Spel_Genre WHERE bgSpel = " & CLng(lID) & ";DELETE FROM cms_Bind_Spel_Spelserie WHERE bsSpel = " & CLng(lID))
          
          For Each nx In aGenre
            If Not IsNumeric(nx) Or nx = Empty Then nx = 0
            If nx > 0 Then sqlInserts = sqlInserts & "INSERT INTO cms_Bind_Spel_Genre (bgGenre, bgSpel) VALUES(" & CLng(nx) & "," & CLng(lID) & ")" & ";"
          Next
          
          For Each nx In aGrupp
            If Not IsNumeric(nx) Or nx = Empty Then nx = 0
            If nx > 0 Then sqlInserts = sqlInserts & "INSERT INTO cms_Bind_Spel_Spelserie (bsSpelSerie, bsSpel) VALUES(" & CLng(nx) & "," & CLng(lID) & ")" & ";"
          Next
          
          If Len(sqlInserts) > 0 Then
            con.ExeCute(sqlInserts)
          End If
        ' ################
        
        ' #### CLEAN-UP ####
          Con.ExeCute("DELETE FROM cms_Speltitlar WHERE tSpelID = 0 AND tSparadAv = " & CLng(cCMS_ID))
        ' ##################
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
    
    'Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
  Case "del" ' Radera
    If GetAcc("CMS44") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM cms_Spel WHERE sID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              Call AddLogg("SPEL","RADERA [TOTAL]",rsDB("sID"))
              
              Set rsBilder = Server.CreateObject("ADODB.RecordSet")
              SQL = "SELECT * FROM cms_Bind_Spel_Img WHERE bsSpel = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("bsBild")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close

              SQL = "SELECT * FROM cms_Speltitlar WHERE tSpelID = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("tBoxart_BoxFram")
                  ImgRemove rsBilder("tBoxart_BoxBak")
                  ImgRemove rsBilder("tBoxart_Manual")
                  ImgRemove rsBilder("tBoxart_Kassett")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close
              Set rsBilder = Nothing
              
              Con.ExeCute("DELETE FROM cms_Speltitlar WHERE tSpelID = " & CLng(oID))
              Con.ExeCute("DELETE FROM cms_Bind_Spel_Genre WHERE bgSpel = " & CLng(oID))
              Con.ExeCute("DELETE FROM cms_Bind_Spel_Spelserie WHERE bsSpel = " & CLng(oID))
              Con.ExeCute("DELETE FROM cms_Bind_Spel_Img WHERE bsSpel = " & CLng(oID))
              rsDB.Delete
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade spel som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Databas/Spel/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>