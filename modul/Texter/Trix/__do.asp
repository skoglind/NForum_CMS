<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS111") Then Response.Redirect("/")
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
    lSpelID     = mForm("vSpelID", "123", 0)
    sTitel      = mForm("vTitel", "ABC", 100)
    sTextM      = mForm("vText", "ABC", 0)
    
    Con_Open
      'If NOT GetAcc("CMS111") Then sbFilter = " AND NOT rStatus = 0"
    
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Speltrix WHERE xID = " & CLng(lID) & sbFilter
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If lSpelID = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du valt inte har valt ett spel.</p>"
        If Len(sTitel) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon titel.</p>"
        
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
          rsDB("xDatumSkapad")  = Now
          rsDB("xSkapadAv")     = cCMS_ID
          rsDB("xStatus")       = 2
        Else
          ' #### ÄNDRA STATUS ####
          lStatus     = mForm("vStatus_cp", "123", 0)
          lPublTyp    = mForm("vPublTyp_cp", "123", 0)
          'lPublDatum  = mForm("vPublDatum_cp", "DAT", 0)
          
          nStatus = rsDB("xStatus")
          
          Select Case lStatus
            Case 2  ' Inväntar publicering
              nStatus = 2
              Call AddLogg("SPELTRIX","PUBLICERING [INVÄNTA]",lID)
            Case 3  ' Publicering nekad
              nStatus = 3
              Call AddLogg("SPELTRIX","PUBLICERING [NEKAD]",lID)
            Case 4  ' Publicerad
              nStatus = 4
              
              ' #### PUBLICERA, VÄLJ DATUM ####
              Select Case lPublTyp
                Case 1  ' DIREKT
                  rsDB("xDatumPublicerad") = Now
                  Call AddLogg("SPELTRIX","PUBLICERING [GODKÄNN DIREKT]",lID)
                Case 2  ' IMORGON
                  rsDB("xDatumPublicerad") = CDate(DateAdd("d", 1, Date) & " " & PUBL_TID)
                  Call AddLogg("SPELTRIX","PUBLICERING [GODKÄNN IMORGON]",lID)
                Case 3  ' NÄSTA VARDAG
                  Do Until IsVardag
                    dPublicerad = DateAdd("d", 1, Date)
                    If weekday(dPublicerad) <> 6 And weekday(dPublicerad) <> 7 Then IsVardag = True
                  Loop
                  rsDB("xDatumPublicerad") = CDate(dPublicerad & " " & PUBL_TID)
                  Call AddLogg("SPELTRIX","PUBLICERING [GODKÄNN VARDAG]",lID)
                Case 4  ' VALT DATUM
                  If Not lPublDatum > #2025-01-01# Then 
                    rsDB("xDatumPublicerad") = CDate(lPublDatum & " " & PUBL_TID)
                    Call AddLogg("SPELTRIX","PUBLICERING [GODKÄNN DATUM]",lID)
                  Else
                    bDontChange = True
                  End If
                Case Else
                  bDontChange = True
              End Select
               
              If nStatus = 4 And Not bDontChange Then rsDB("xPubliceradAv") = cCMS_ID
              ' ###############################
              
            Case Else
              bDontChange = True
          End Select
          
          If Not bDontChange Then rsDB("xStatus") = nStatus
          ' ######################
        End If
        
        rsDB("xSpelID")   = lSpelID
        rsDB("xTitel")    = sTitel
        rsDB("xTextM")    = sTextM
        
        rsDB("xDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("xID")
        If bIsNew Then AddLogg "SPELTRIX","SKAPA",lID
        
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
    Con_Open
      allID = Split(GetFormRequest("chk_id", "YES"), ",")
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      
        For Each oID IN allID
          'If Not GetAcc("CMS111") Then sbFilter = " AND rSkapadAv = " & cCMS_ID
          SQL = "SELECT * FROM cms_Speltrix WHERE xID = " & CLng(oID)
          rsDB.Open SQL, Con, 1, 3
        
          If Not rsDB.Eof Then
            If GetAcc("CMS111") Then
              Call AddLogg("SPELTRIX","RADERA [TOTAL]",rsDB("xID"))
              
              rsDB.Delete
            End If
          End If
          
          rsDB.Close
        Next
      
      Set rsDB = Nothing
    Con_Close
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade tips & trixen som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/Texter/Trix/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>