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
    sNyckelord  = mForm("vNyckelord", "ABC", 500)
    sTextM      = mForm("vText", "ABC", 0)
    
    sSynlig     = mForm("vSynlig_cp", "CHK", 0)
    
    nID         = FormIDToArray("vRegion_")
    nUseTitel   = mForm("vStandardTitel", "ABC", 50)
    
    ReDim nRegion(UBound(nID))
    ReDim nSortNo(UBound(nID))
    ReDim nTitel(UBound(nID))
    ReDim nRelease(UBound(nID))
    ReDim nRegionskod(UBound(nID))
    ReDim nAction(UBound(nID))
    
    For zx = 0 To UBound(nID)
      nID(zx) = Trim(nID(zx))
    
      nRegion(zx)     = mForm("vRegion_" & nID(zx), "123", 0)
      nSortNo(zx)     = mForm("vSortNo_" & nID(zx), "123", 0)
      nTitel(zx)      = mForm("vTitel_" & nID(zx), "ABC", 255)
      nRelease(zx)    = mForm("vRelease_" & nID(zx), "ABC", 40)
      nRegionskod(zx) = mForm("vRegionskod_" & nID(zx), "ABC", 50)
      nAction(zx)     = False
    Next
     
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Tillbehor WHERE iID = " & CLng(lID)
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
        
        rsDB("iKonsol")       = CLng(sKonsol)
        rsDB("iNyckelord")    = sNyckelord
        rsDB("iTextM")        = sTextM
        
        rsDB("iSynlig")       = sSynlig
        
        rsDB("iDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
        
        lID = rsDB("iID")
        If bIsNew Then AddLogg "TILLBEHÖR","SKAPA",lID
        
        ' #### TITLAR ####
          Set rsT = Server.CreateObject("ADODB.RecordSet")
          SQL = "SELECT * FROM cms_Tillbehortitlar WHERE tTillbehorID = " & CLng(lID) & " ORDER BY tID ASC"
          rsT.Open SQL, Con, 1, 3
          
          lFindID = -1
          Do Until rsT.EOF
            lFindID = IsInArray(nID, rsT("tID"))
            If lFindID < 0 Then lFindID = IsInArray(nID, rsT("tSparadNyckel"))
            
            If lFindID < 0 Then
              ImgRemove rsT("tBoxart_BoxFram")
              ImgRemove rsT("tBoxart_BoxBak")
              ImgRemove rsT("tBoxart_Manual")
              ImgRemove rsT("tBoxart_Tillbehor")
            
              rsT.Delete
            Else
              nAction(lFindID)    = True
              rsT("tTitel")       = Left(Trim(nTitel(lFindID) & " "),255)
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
              SQL = "SELECT * FROM cms_Tillbehortitlar WHERE tSparadNyckel = '" & MakeLegal(nID(zx)) & "' AND tSparadAv = " & CLng(cCMS_ID)
              rsT.Open SQL, Con
              
              If rsT.EOF Then
                rsT.AddNew
              End If
              
              nAction(zx)           = True
              rsT("tTillbehorID")        = lID
              rsT("tTitel")         = Left(Trim(nTitel(zx)),255)
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
          
          Con.ExeCute("UPDATE cms_Tillbehor SET iStandard_Titel = " & CLng(nUseTitel) & " WHERE iID = " & CLng(lID))
        ' ################
        
        ' #### CLEAN-UP ####
          Con.ExeCute("DELETE FROM cms_Tillbehortitlar WHERE tTillbehorID = 0 AND tSparadAv = " & CLng(cCMS_ID))
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
            SQL = "SELECT * FROM cms_Tillbehor WHERE iID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              Call AddLogg("TILLBEHÖR","RADERA [TOTAL]",rsDB("iID"))

              Set rsBilder = Server.CreateObject("ADODB.RecordSet")
              SQL = "SELECT * FROM cms_Tillbehortitlar WHERE tTillbehorID = " & CLng(oID)
              rsBilder.Open SQL, Con
              
                Do Until rsBilder.EOF
                  ImgRemove rsBilder("tBoxart_BoxFram")
                  ImgRemove rsBilder("tBoxart_BoxBak")
                  ImgRemove rsBilder("tBoxart_Manual")
                  ImgRemove rsBilder("tBoxart_Tillbehor")
                  rsBilder.MoveNext
                Loop
              
              rsBilder.Close
              Set rsBilder = Nothing
              
              Con.ExeCute("DELETE FROM cms_Tillbehortitlar WHERE tTillbehorID = " & CLng(oID))

              rsDB.Delete
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade tillbehör som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Databas/Tillbehor/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>