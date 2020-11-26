<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS2") Then Response.Redirect("/")
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
    lID         = mForm("vID", "123", 0)
    lTitel      = mForm("vTitel", "123", 0)
    sEgenTitel  = mForm("vEgenTitel", "ABC", 20)
    sAnvNamn    = mForm("vAnvNamn", "ABC", 60)
    sEpost      = mForm("vEpost", "ABC", 255)
    sNamn       = mForm("vNamn", "ABC", 50)
    sPlats      = mForm("vPlats", "ABC", 50)
    sHemsida    = mForm("vHemsida", "ABC", 255)
    sMSN        = mForm("vMSN", "ABC", 255)
    sICQ        = mForm("vICQ", "ABC", 255)
    sSignatur   = mForm("vSignatur", "ABC", 255)
    sProfil     = mForm("vText", "ABC", 0)
    
    sLosen1     = mForm("vLosen1", "ABC", 50)
    sLosen2     = mForm("vLosen2", "ABC", 50)
    
    bBannad     = mForm("vBannad_cp", "123", 0)
    
    bCMS        = mForm("vCMS_cp", "CHK", 0)
    
    bCMS0       = mForm("vCMS0_cp", "123", 0)
    bCMS1       = mForm("vCMS1_cp", "123", 0)
    bCMS2       = mForm("vCMS2_cp", "123", 0)
    bCMS3       = mForm("vCMS3_cp", "123", 0)
    bCMS4       = mForm("vCMS4_cp", "123", 0)
    bCMS5       = mForm("vCMS5_cp", "123", 0)
    bCMS6       = mForm("vCMS6_cp", "123", 0)
    bCMS7       = mForm("vCMS7_cp", "123", 0)
    
    bAktiverad  = mForm("vAktiverad_cp", "CHK", 0)
    
    Con_Open
      If NOT GetAcc("CMS202") Then sbFilter = ""
    
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID) & sbFilter
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sLosen1) > 0 Then If Len(sLosen1) < 5 Or sLosen1 <> sLosen2 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du angav ett för kort lösenord eller så stämmer de inte överrens.</p>"
        If Len(sNamn) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett något namn.</p>"
        If GetAcc("CMS202") Then If Len(sEpost) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon e-postadress.</p>"
        If GetAcc("CMS202") Then If lTitel = 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du valt en otillåten titel.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
        Else
          bIsNew = True
          If GetAcc("CMS202") Then If Len(sLosen1) < 5 Or sLosen1 <> sLosen2 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du angav ett för kort lösenord eller så stämmer de inte överrens.</p>"
          If GetAcc("CMS202") Then If MakeLegal(sAnvNamn) <> sAnvNamn Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då användarnamnet är ogiltigt.</p>"
          If GetAcc("CMS202") Then If Not DoUserExist(sAnvNamn) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då användarnamnet redan används.</p>"
          If GetAcc("CMS202") Then If Len(sAnvNamn) < 5 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då användarnamnet är för kort.</p>"
          If Not GetAcc("CMS202") Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har behörighet att skapa användare.</p>"
        End If
        
        If bErr Then
          Response.Write("<script type='text/javascript'>parent.savefailed('" & nMessage & "');</script>")
          Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
          Response.End
        End If
      ' ######################
      
        If rsDB.EOF Then
          rsDB.AddNew
          rsDB("aAnvNamn")        = MakeLegal(sAnvNamn)
          rsDB("aBlockadTill")    = #2003-01-01 00:00:00#
          rsDB("aMedlemSedan")    = Now
          rsDB("aTimeStamp")      = DateAdd("n", -10, Now)
          rsDB("aInloggadSenast") = DateAdd("n", -10, Now)
          ' ## GBDB EXKLUSIVT ##
          rsDB("aNewActivation") = True
          rsDB("aNewDelivered")  = True
          ' ####################
        End If
        
        If GetAcc("CMS202") Then
          rsDB("aTitelID")    = lTitel
          rsDB("aEpost")      = sEpost
          
          If Len(sLosen1) > 0 Then
            rsDB("aNyttLosenord")  = True
            nSalt1                 = SlumpText(5)
            nSalt2                 = SlumpText(5)
            rsDB("aSalt1")         = nSalt1
            rsDB("aSalt2")         = nSalt2
            rsDB("aPassWd")        = MD5(config_Hash_Salt_1 & "" & nSalt1 & "" & sLosen1 & "" & config_Hash_Salt_2 & "" & nSalt2)
          End If
        End If
        
        rsDB("aEgenTitel")  = sEgenTitel
        rsDB("aNamn")       = sNamn
        rsDB("aPlats")      = sPlats
        rsDB("aHemsida")    = sHemsida
        rsDB("aMSN")        = sMSN
        rsDB("aICQ")        = sICQ
        rsDB("aSignatur")   = sSignatur
        rsDB("aPM")         = sProfil
        
        If CLng(lID) <> CLng(cCMS_ID) Then
          If CLng(bBannad) <> CLng(rsDB("aBlockStatus")) Then
            Select Case bBannad
              Case 1
                rsDB("aBlockadTill") = DateAdd("d", 7, Now)
                rsDB("aBlockStatus") = 1
              Case 2  
                rsDB("aBlockadTill") = DateAdd("m", 1, Now)
                rsDB("aBlockStatus") = 2
              Case 3  
                rsDB("aBlockadTill") = DateAdd("m", 3, Now)
                rsDB("aBlockStatus") = 3
              Case 4 
                rsDB("aBlockadTill") = DateAdd("m", 6, Now)
                rsDB("aBlockStatus") = 4
              Case 5
                rsDB("aBlockadTill") = DateAdd("yyyy", 1, Now)
                rsDB("aBlockStatus") = 5
              Case 6 
                rsDB("aBlockadTill") = DateAdd("yyyy", 25, Now)
                rsDB("aBlockStatus") = 6
              Case Else
                rsDB("aBlockadTill") = #2003-01-01 00:00:00#
                rsDB("aBlockStatus") = 0
            End Select
          End If
        End If
        
        If GetAcc("CMS202") Then
          If CLng(lID) <> CLng(cCMS_ID) Then rsDB("aAktiverad") = bAktiverad
        
          rsDB("aS_CMS") = bCMS
          
          secString = ""
          Select Case bCMS0
            Case 1 : secString = secString & "CMS000;"
          End Select
          
          Select Case bCMS1
            Case 1 : secString = secString & "CMS100;"
            Case 2 : secString = secString & "CMS110;"
            Case 3 : secString = secString & "CMS111;"
          End Select
          
          If CLng(lID) <> CLng(cCMS_ID) Then
            Select Case bCMS2
              Case 1 : secString = secString & "CMS200;"
              Case 2 : secString = secString & "CMS202;"
            End Select
          Else
            secString = secString & "CMS202;"
          End If
          
          Select Case bCMS3
            Case 1 : secString = secString & "CMS300;"
            Case 2 : secString = secString & "CMS330;"
            Case 3 : secString = secString & "CMS333;"
          End Select
          
          Select Case bCMS4
            Case 1 : secString = secString & "CMS400;"
            Case 2 : secString = secString & "CMS440;"
            Case 3 : secString = secString & "CMS444;"
          End Select
          
          Select Case bCMS5
            Case 1 : secString = secString & "CMS500;"
          End Select
          
          Select Case bCMS6
            Case 1 : secString = secString & "CMS600;"
          End Select
          
          Select Case bCMS7
            Case 1 : secString = secString & "CMS700;"
          End Select
          
          If secString = Empty Then 
            secString = "0"
          Else
            secString = Left(secString, Len(secString)-1)
          End If
          
          rsDB("aS_CMSRatter") = secString
        End If
        
        rsDB("aDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("aID")
        If bIsNew Then AddLogg "ANVÄNDARE","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
    Con_Close
    
    Select Case sExtraAction
      Case "return"
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",true,'" & "_show.asp?" & sRebuild & "');</script>")
      Case Else
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",false,'');</script>")
    End Select
    
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
  Case "logout"
    If GetAcc("CMS202") Then
      Con_Open
        lID = mForm("vID", "123", 0)
        Con.ExeCute("UPDATE fsBB_Anv SET aLOCK = 1 WHERE aID = " & CLng(lID))
      Con_Close
    End If
    Response.Write("<script type='text/javascript'>parent.dofinished('Användaren är nu utloggad!');</script>")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>