<%
Server.ScriptTimeout = 600

Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.CacheControl = "no-cache"
%>

<!--#INCLUDE FILE="cms_Config.asp"-->
<!--#INCLUDE FILE="cms_Constant.asp"-->
<!--#INCLUDE FILE="cms_Functions.asp"-->
<!--#INCLUDE FILE="cms_Lists.asp"-->
<!--#INCLUDE FILE="cms_UplSecurity.asp"-->

<%
lMode = Request.QueryString("m")
If Not IsNumeric(lMode) Or lMode = Empty Then lMode = 0
lMode = CLng(lMode)

Select Case lMode
  Case 1 ' BIG SIZE UPLOAD
    On Error Resume Next
    
    Con_Open
    Set Upl = Server.CreateObject("aspSmartUpload.SmartUpload")
    
      Upl.MaxFileSize       = UPLOAD_MAXSIZE
      Upl.AllowedFilesList  = "jpg,jpeg,png,bmp,gif"
      Upl.Upload
      
      Set File = Upl.Files.Item(1)
      
        lID   = FixGet(upl.Form("f_id"),"123", 23)
        objID = FixGet(upl.Form("f_objid"),"123", 0) 
        sArea = LCase(FixGet(upl.Form("f_area"), "ABC", 25))
      
        Select Case sArea
          Case "news"
            If GetAcc("CMS1") Then
              dbTable = "cms_Bind_Nyheter_Img" : dbUser = "bnUser" : dbBild = "bnBild" : dbSaved = "bnSaved" : dbBindField = "bnNyhet" : dbAllFields = "bnBildText,ABC,250"
              If Not IsBehorig("NEWS", objID) Then anyError = True
            Else
              anyError = True
            End If
          Case "game"
            If GetAcc("CMS4") Then
              dbTable = "cms_Bind_Spel_Img" : dbUser = "bsUser" : dbBild = "bsBild" : dbSaved = "bsSaved" : dbBindField = "bsSpel" : dbAllFields = "bsBildText,ABC,250"
              If Not IsBehorig("GAME", objID) Then anyError = True
            Else
              anyError = True
            End If
          Case "rec"
            If GetAcc("CMS1") Then
              dbTable = "cms_Bind_Recension_Img" : dbUser = "brUser" : dbBild = "brBild" : dbSaved = "brSaved" : dbBindField = "brRecension" : dbAllFields = "brBildText,ABC,250"
              If Not IsBehorig("REC", objID) Then anyError = True
            Else
              anyError = True
            End If
          Case "art"
            If GetAcc("CMS1") Then
              dbTable = "cms_Bind_Artikel_Img" : dbUser = "baUser" : dbBild = "baBild" : dbSaved = "baSaved" : dbBindField = "baArtikel" : dbAllFields = "baBildText,ABC,250"
              If Not IsBehorig("ART", objID) Then anyError = True
            Else
              anyError = True
            End If
          Case Else
            anyError = True
        End Select
        
        Set rsDB = Server.CreateObject("ADODB.Recordset")
        
          If Not anyError Then
            ' ## KOLLA OM EN BILD FINNS DÄR ##
            bFileUploaded = True
            If Err.Number = -2147220399 Or Err.Number = -2147220299 Or File.IsMissing Then bFileUploaded = False
            ' ################################
          
            ' ## SKAPA BILDENS LÄNK I DATABASEN ##
            If bFileUploaded Then
              sOriginalNamn = LCase(Left(Trim(File.Filename), 500))
              sFilandelse   = LCase(Left(Trim(File.Fileext),25))
            
              SQL = "SELECT * FROM cms_Bild WHERE bID = " & lID
              rsDB.Open SQL, Con, 1, 3
              
                If rsDB.EOF Then 
                  rsDB.AddNew
                  bIsNew = True
                  rsDB("bSparad") = True
                Else
                  ImgRemove rsDB("bID")
                End If
                
                rsDB("bUppladdadAv")  = cCMS_ID
                rsDB("bOriginalNamn") = sOriginalNamn
                rsDB("bTyp")          = sFilandelse
                rsDB("bInSizes")      = ";"
              
                rsDB.Update
                lID = rsDB("bID")
                
                sFilnamn      = "img_" & Right("0000000000" & lID, 10) & "_original." & File.Fileext
                File.SaveAs UPLOAD_FAKE & sFilnamn
                
              rsDB.Close
            Else
              ' Har man skapat en post måste så klart en bild laddas upp...
              cID = Con.ExeCute("SELECT COUNT(*) FROM cms_Bild WHERE bID = " & lID)(0)
              If Not IsNumeric(cID) Or cID = Empty Then cID = 0
              If cID = 0 Then anyError = True
            End If
            ' ####################################
          End If
          
          ' ## BIND BILDEN MED RÄTT DATABAS ##
          If Not anyError Then
            SQL = "SELECT * FROM " & dbTable & " WHERE " & dbBild & " = " & lID
            rsDB.Open SQL, Con, 1, 3
            
              If rsDB.EOF Then
                rsDB.AddNew
                rsDB(dbBild)    = lID
                rsDB(dbBindField) = CLng(objID)
                If CLng(objID) = 0 Then rsDB(dbSaved) = False Else rsDB(dbSaved) = True
              End If
              
              rsDB(dbUser)     = cCMS_ID
              
              dbFields = Split(dbAllFields, ":")
              For Each dbField In dbFields
                nF = Split(dbField, ",")
                rsDB(nF(0)) = FixGet(upl.Form("f_" & nF(0)), nF(1), CLng(nF(2)))
                
                sendBack = sendBack & "tetra_f_" & nF(0) & "||" & rsDB(nF(0)) & "||"
              Next
              
              rsDB.Update
            
            rsDB.Close
          End If
          ' ##################################
          
          ' ## FIXA FILNAMNET ##
          If Not anyError Then
            cAndelse = Con.ExeCute("SELECT bTyp FROM cms_Bild WHERE bID = " & lID)(0)
            sFilnamn = "/cms_img.asp?e=" & lID & "&w=80&h=80"
          End If
          ' ####################
          
          Set File = Nothing
        Set rsDB = Nothing
      Set Upl = Nothing
      Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.uploadfailed();</script>")
    Else
      If bIsNew Then bIsNew = "true" Else bIsNew = "false"
      Response.Write("<script type='text/javascript'>parent.uploadfinished(" & bIsNew & "," & lID & ",'" & sFilnamn & "','" & sendBack & "');</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case 2 ' BIG SIZE UPLOAD - DELETE
    Con_Open
      lID   = mGet("vID", "123", 0)
      objID = mGet("vObjID", "123", 0)
      sArea = mGet("vArea", "ABC", 25)
      
      anyError = False
      Select Case LCase(sArea)
        Case "news"
          If GetAcc("CMS1") Then
            dbTable = "cms_Bind_Nyheter_Img" : dbBindField = "bnNyhet" : dbBindBild = "bnBild" : dbBindID = "bnID"
            If Not IsBehorig("NEWS", objID) Then anyError = True
          Else
            anyError = True
          End If
        Case "game"
          If GetAcc("CMS4") Then
            dbTable = "cms_Bind_Spel_Img" : dbBindField = "bsSpel" : dbBindBild = "bsBild" : dbBindID = "bsID"
            If Not IsBehorig("GAME", objID) Then anyError = True
          Else
            anyError = True
          End If
        Case "rec"
          If GetAcc("CMS1") Then
            dbTable = "cms_Bind_Recension_Img" : dbBindField = "brRecension" : dbBindBild = "brBild" : dbBindID = "brID"
            If Not IsBehorig("REC", objID) Then anyError = True
          Else
            anyError = True
          End If
        Case "art"
          If GetAcc("CMS1") Then
            dbTable = "cms_Bind_Artikel_Img" : dbBindField = "baArtikel" : dbBindBild = "baBild" : dbBindID = "baID"
            If Not IsBehorig("ART", objID) Then anyError = True
          Else
            anyError = True
          End If
        Case Else
          anyError = True
      End Select
      
      If Not anyError Then
        Set rsBild = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM " & dbTable & " LEFT JOIN cms_Bild ON " & dbTable & "." & dbBindBIld & " = cms_Bild.bID WHERE " & dbBindField & " = " & CLng(objID) & " AND " & dbBindBIld & " = " & CLng(lID)
        rsBild.Open SQL, Con
        
          If rsBild.EOF Then
            anyError = True
          Else
            rmBildID  = rsBild("bID")
            rmBindID  = rsBild(dbBindID)
          End If
          
        rsBild.Close
        Set rsBild = Nothing
        
        If Not anyError Then
          Con.ExeCute("DELETE FROM " & dbTable & " WHERE " & dbBindID & " = " & CLng(rmBindID))
          
          ImgRemove rmBildID
        End If
      End If
    Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.deletefailed();</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.deletefinished(" & lID & ");</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case 3 ' SMALL SIZE UPLOAD
    On Error Resume Next
    
    Con_Open
    Set Upl = Server.CreateObject("aspSmartUpload.SmartUpload")
    
      Upl.MaxFileSize       = UPLOAD_MAXSIZE
      Upl.AllowedFilesList  = "jpg,jpeg,png,bmp,gif"
      Upl.Upload
      
      Set File = Upl.Files.Item(1)
      
        dataID = FixGet(upl.Form("lID"),"123", 23)
        sArea  = FixGet(upl.Form("sArea"),"ABC", 20)
      
        Select Case sArea
          Case "foretag"
            If GetAcc("CMS4") Then
              dbTable = "cms_Foretag" : dbBindField = "fLogga" : dbTableID = "fID"
            Else
              anyError = True
            End If
          Case "nyheter"
            If GetAcc("CMS4") Then
              dbTable = "cms_Nyheter" : dbBindField = "nFlash" : dbTableID = "nID"
            Else
              anyError = True
            End If
          Case "artiklar"
            If GetAcc("CMS4") Then
              dbTable = "cms_Artiklar" : dbBindField = "aaFlash" : dbTableID = "aaID"
            Else
              anyError = True
            End If
          Case "recensioner"
            If GetAcc("CMS4") Then
              dbTable = "cms_Recensioner" : dbBindField = "rFlash" : dbTableID = "rID"
            Else
              anyError = True
            End If
          Case Else
            anyError = True
        End Select

        Set rsData = Server.CreateObject("ADODB.Recordset")
          SQL = "SELECT * FROM " & dbTable & " WHERE " & dbTableID & " = " & CLng(dataID)
          rsData.Open SQL, COn
            If rsData.EOF Then
              lID = 0
              bSaved = False
              eBild = GetImgIDByExclusive()
              If eBild > 0 Then
                bSaved = True
                lID = eBild
              End If
            Else
              lID = rsData(dbBindField)
              bSaved = True
            End If
          rsData.Close
        Set rsData = Nothing
        
        Set rsDB = Server.CreateObject("ADODB.Recordset")

          If Not anyError Then
            ' ## KOLLA OM EN BILD FINNS DÄR ##
            If Err.Number = -2147220399 Or Err.Number = -2147220299 Or File.IsMissing Then anyError = True
            ' ################################
          
            ' ## SKAPA BILDENS LÄNK I DATABASEN ##
            If Not anyError Then
              sOriginalNamn = LCase(Left(Trim(File.Filename), 500))
              sFilandelse   = LCase(Left(Trim(File.Fileext),25))
            
              SQL = "SELECT * FROM cms_Bild WHERE bID = " & lID
              rsDB.Open SQL, Con, 1, 3
              
                If rsDB.EOF Then
                  rsDB.AddNew
                  rsDB("bSparad") = bSaved
                  If Not bSaved Then rsDB("bExclusiveID") = Make_ExclusiveKey
                Else
                  ImgRemove rsDB("bID")
                ENd If
                
                rsDB("bUppladdadAv")  = cCMS_ID
                rsDB("bOriginalNamn") = sOriginalNamn
                rsDB("bTyp")          = sFilandelse
                rsDB("bInSizes")      = ";"
              
                rsDB.Update
                lID = rsDB("bID")
                
                sFilnamn  = "img_" & Right("0000000000" & lID, 10) & "_original." & File.Fileext
                File.SaveAs UPLOAD_FAKE & sFilnamn
                
              rsDB.Close
              
              Con.Execute("UPDATE " & dbTable & " SET " & dbBindField & " = " & CLng(lID) & " WHERE " & dbTableID & " = " & CLng(dataID))
            End If
            ' ####################################
          End If
          
          Set File = Nothing
        Set rsDB = Nothing
      Set Upl = Nothing
      Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.smalluploadfailed();</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.smalluploadfinished(" & CLng(dataID) & ");</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case 4 ' SMALL SIZE UPLOAD - DELETE
    Con_Open
      dataID = mGet("lID", "123", 0)
      sArea  = mGet("sArea", "ABC", 25)
      
      anyError = False
      Select Case LCase(sArea)
        Case "foretag"
          If GetAcc("CMS4") Then
            dbTable = "cms_Foretag" : dbBindField = "fLogga" : dbTableID = "fID"
          Else
            anyError = True
          End If
        Case "nyheter"
          If GetAcc("CMS4") Then
            dbTable = "cms_Nyheter" : dbBindField = "nFlash" : dbTableID = "nID"
          Else
            anyError = True
          End If
        Case "artiklar"
          If GetAcc("CMS4") Then
            dbTable = "cms_Artiklar" : dbBindField = "aaFlash" : dbTableID = "aaID"
          Else
            anyError = True
          End If
        Case "recensioner"
          If GetAcc("CMS4") Then
            dbTable = "cms_Recensioner" : dbBindField = "rFlash" : dbTableID = "rID"
          Else
            anyError = True
          End If
        Case Else
          anyError = True
      End Select
      
      If Not anyError Then
        Set rsBild = Server.CreateObject("ADODB.RecordSet")
          SQL = "SELECT * FROM " & dbTable & " LEFT JOIN cms_Bild ON " & dbTable & "." & dbBindField & " = cms_Bild.bID WHERE " & dbTableID & " = " & CLng(dataID)
          rsBild.Open SQL, Con, 1, 3
          
            If rsBild.EOF Then
              anyError = True
              dataID = 0
            Else
              dataID = rsBild(dbTableID)
              If IsNull(rsBild("bID")) Then 
                anyError = True
              Else
                rmBildID  = CLng(rsBild("bID"))
                
                rsBild(dbBindField) = 0
                rsBild.Update
              End If
            End IF
          
          rsBild.Close
        Set rsBild = Nothing
        
        If Not anyError Then
          ImgRemove rmBildID
        End If
      End If
    Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.smalldeletefailed();</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.smalldeletefinished(" & dataID & ");</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case 5 ' BOXART UPLOAD
    On Error Resume Next
    
    Con_Open
    Set Upl = Server.CreateObject("aspSmartUpload.SmartUpload")
    
      Upl.MaxFileSize       = UPLOAD_MAXSIZE
      Upl.AllowedFilesList  = "jpg,jpeg,png,bmp,gif"
      Upl.Upload
      
      Set File = Upl.Files.Item("upload_Boxart")
      
        dataID = FixGet(upl.Form("uID"),"ABC", 50)
        sArt   = FixGet(upl.Form("uArt"),"123", 0)
        gameID = FixGet(upl.Form("uGameID"),"123", 0)
        sArea  = FixGet(upl.Form("uArea"),"ABC", 10)
      
        If Not IsNumeric(dataID) Or dataID = Empty Then 
          transID = 0
        Else
          transID = CLng(dataID)
        End If
        
        Select Case sArea
          Case "game"
            If GetAcc("CMS4") Then
              dbTable = "cms_Speltitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel" : dbSpelID = "tSpelID"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Kassett"
                Case Else : anyError = True
              End Select
            Else
              anyError = True
            End If
          Case "console"
            If GetAcc("CMS4") Then
              dbTable = "cms_Konsoltitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel" : dbSpelID = "tKonsolID"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Konsol"
                Case Else : anyError = True
              End Select
            Else
              anyError = True
            End If
          Case "addon"
            If GetAcc("CMS4") Then
              dbTable = "cms_Tillbehortitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel" : dbSpelID = "tTillbehorID"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Tillbehor"
                Case Else : anyError = True
              End Select
            Else
              anyError = True
            End If
          Case Else
            anyError = True
        End Select

        Set rsData = Server.CreateObject("ADODB.Recordset")
          SQL = "SELECT * FROM " & dbTable & " WHERE " & dbID & " = " & CLng(transID) & " OR (" & dbUser & " = " & CLng(cCMS_ID) & " AND " & dbKey & " = '" & Trim(MakeLegal(dataID)) & "')"
          rsData.Open SQL, Con, 1, 3
            If rsData.EOF Then
              rsData.AddNew
                rsData(dbKey)  = Trim(MakeLegal(dataID))
                rsData(dbUser) = CLng(cCMS_ID)
                rsData(dbTitel) = "-- INGEN TITEL ANGIVEN --"
                rsData(dbSpelID) = gameID
              rsData.Update
              
              lID = 0
            Else
              lID = rsData(dbBindField)
            End If

            titelID = rsData(dbID)
          rsData.Close
        Set rsData = Nothing
        
        Set rsDB = Server.CreateObject("ADODB.Recordset")

          If Not anyError Then
            ' ## KOLLA OM EN BILD FINNS DÄR ##
            If Err.Number = -2147220399 Or Err.Number = -2147220299 Or File.IsMissing Then anyError = True
            ' ################################
          
            ' ## SKAPA BILDENS LÄNK I DATABASEN ##
            If Not anyError Then
              sOriginalNamn = LCase(Left(Trim(File.Filename), 500))
              sFilandelse   = LCase(Left(Trim(File.Fileext),25))
            
              SQL = "SELECT * FROM cms_Bild WHERE bID = " & lID
              rsDB.Open SQL, Con, 1, 3
              
                If rsDB.EOF Then
                  rsDB.AddNew
                  rsDB("bSparad") = True
                ENd If
                
                rsDB("bUppladdadAv")  = cCMS_ID
                rsDB("bOriginalNamn") = sOriginalNamn
                rsDB("bTyp")          = sFilandelse
                rsDB("bInSizes")      = ";"
              
                rsDB.Update
                lID = rsDB("bID")
                
                sFilnamn      = "img_" & Right("0000000000" & lID, 10) & "_original." & File.Fileext
                sFilnamnSend  = "/cms_Img.asp?e=" & lID & "&w=80&h=80"
                File.SaveAs UPLOAD_FAKE & sFilnamn
                
              rsDB.Close
              
              Con.Execute("UPDATE " & dbTable & " SET " & dbBindField & " = " & CLng(lID) & " WHERE " & dbID & " = " & CLng(titelID))
            End If
            ' ####################################
          End If
          
          Set File = Nothing
        Set rsDB = Nothing
      Set Upl = Nothing
      Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.boxuploadfailed();</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.boxuploadfinished('" & dataID & "'," & sArt & ",'" & sFilnamnSend & "');</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case 6 ' BOXART UPLOAD - DELETE
    Con_Open
      dataID = mGet("lID", "ABC", 50)
      sArt   = mGet("sArt", "123", 25)
      sArea  = mGet("sArea", "ABC", 25)
      
      If Not IsNumeric(dataID) Or dataID = Empty Then 
        transID = 0
      Else
        transID = CLng(dataID)
      End If
      
      anyError = False
      Select Case LCase(sArea)
        Case "game"
          If GetAcc("CMS4") Then
            dbTable = "cms_Speltitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Kassett"
                Case Else : anyError = True
              End Select
          Else
            anyError = True
          End If
        Case "console"
          If GetAcc("CMS4") Then
            dbTable = "cms_Konsoltitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Konsol"
                Case Else : anyError = True
              End Select
          Else
            anyError = True
          End If
        Case "addon"
          If GetAcc("CMS4") Then
            dbTable = "cms_Tillbehortitlar" : dbID = "tID" : dbKey = "tSparadNyckel" : dbUser = "tSparadAv" : dbTitel = "tTitel"
              Select Case CLng(sArt)
                Case 1    : dbBindField = "tBoxart_BoxFram"
                Case 2    : dbBindField = "tBoxart_BoxBak"
                Case 3    : dbBindField = "tBoxart_Manual"
                Case 4    : dbBindField = "tBoxart_Tillbehor"
                Case Else : anyError = True
              End Select
          Else
            anyError = True
          End If
        Case Else
          anyError = True
      End Select
      
      If Not anyError Then
        Set rsBild = Server.CreateObject("ADODB.RecordSet")
          SQL = "SELECT * FROM " & dbTable & " LEFT JOIN cms_Bild ON " & dbTable & "." & dbBindField & " = cms_Bild.bID WHERE " & dbID & " = " & CLng(transID) & " OR (" & dbUser & " = " & CLng(cCMS_ID) & " AND " & dbKey & " = '" & MakeLegal(dataID) & "')"
          rsBild.Open SQL, Con, 1, 3
          
            If rsBild.EOF Then
              anyError = True
            Else
              If IsNull(rsBild("bID")) Then 
                anyError = True
              Else
                rmBildID  = CLng(rsBild("bID"))
                
                rsBild(dbBindField) = 0
                rsBild.Update
              End If
            End IF
          
          rsBild.Close
        Set rsBild = Nothing
        
        If Not anyError Then
          ImgRemove rmBildID
        End If
      End If
    Con_Close
    
    If anyError Then
      Response.Write("<script type='text/javascript'>parent.boxdeletefailed();</script>")
    Else
      Response.Write("<script type='text/javascript'>parent.boxdeletefinished('" & dataID & "'," & sArt & ");</script>")
    End If
    
    Response.Write("<script type='text/javascript'>location.href='_awaiting.asp';</script>")
  Case Else
End Select
%>