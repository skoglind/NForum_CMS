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
    sFraga      = mForm("vFraga", "ABC", 255)
    bSynlig     = mForm("vSynlig", "CHK", 0)
    dSlutDatum  = mForm("vSlutDatum", "ABC", 10)
    
    ' #### DYNAROWS ####
    bSkip = False
    For Each rf In Request.Form
      If Len(rf) > 3 Then If Right(rf, 4) = "XXXX" Then bSkip = True Else bSkip = False
      
      If Not bSkip Then
        sTID = 0
        If LCase(Left(rf,5)) = "vval_" Then
          sTID = Right(rf, Len(rf) - 5)
        
          sID_all = sID_all & ";;" & sTID
          sVal_all = sVal_all & ";;" & mForm("vVal_" & sTID, "ABC", 50)
          sSortNr_all = sSortNr_all & ";;" & mForm("vSortNr_" & sTID, "123", 2)
        End If
      End If
    Next
    
    If Len(sID_all) > 2 Then sID_all = Right(sID_all, Len(sID_all)-2)
    If Len(sVal_all) > 2 Then sVal_all = Right(sVal_all, Len(sVal_all)-2)
    If Len(sSortNr_all) > 2 Then sSortNr_all = Right(sSortNr_all, Len(sSortNr_all)-2)
    
    sID     = Split(sID_all,";;")
    sVal    = Split(sVal_all,";;")
    sSortNr = Split(sSortNr_all,";;")
    ' /#### DYNAROWS ####
    
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_Omrostning_Fraga WHERE omfID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sFraga) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon fråga.</p>"
        If Not IsDate(dSlutDatum) Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du har angett ett ogiltigt slutdatum.</p>"
        
        If UBound(sID) < 0 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du har angett några omröstningsalternativ.</p>"
        
        For zz = 0 To UBound(sVal)
          If Len(sVal(zz)) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon text under minst ett av valen.</p>"
        Next
        
        If Not rsDB.EOF Then
          bIsNew = False
          If Not GetAcc("CMS11") Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du saknar behörighet att ändra denna omröstning.</p>"
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
        
        rsDB("omfFraga")        = sFraga
        rsDB("omfSynlig")       = bSynlig
        rsDB("omfSlutDatum")    = dSlutDatum
        
        rsDB("omfDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("omfID")
        If bIsNew Then AddLogg "OMRÖSTNING","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
      
      ' #### DYNAROWS ####
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
          notDelID = "0"
          For zz = 0 To UBound(sID)
            SQL = "SELECT * FROM cms_Omrostning_Val WHERE omvFraga = " & CLng(lID) & " AND omvID = " & CLng(sID(zz))
            rsDB.Open SQL, Con, 1, 3
            
              If rsDB.EOF Then
                rsDB.AddNew
                rsDB("omvFraga") = lID
              End If
              
              rsDB("omvText")   = sVal(zz)
              rsDB("omvSortNr") = CLng(sSortNr(zz))
              
              rsDB.Update
              notDelID = notDelID & "," & rsDB("omvID")
            rsDB.Close
          Next
        Set rsDB = Nothing
        
        Con.ExeCute("DELETE FROM cms_Omrostning_Val WHERE NOT omvID IN(" & notDelID & ") AND omvFraga = " & CLng(lID))
        Con.ExeCute("DELETE FROM cms_Omrostning_Svar WHERE NOT omsSvar IN(" & notDelID & ") AND omsFraga = " & CLng(lID))
      ' #### /DYNAROWS ####
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
      
        If GetAcc("CMS11") Then
          For Each oID IN allID
            SQL = "SELECT * FROM cms_Omrostning_Fraga WHERE omfID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              Call AddLogg("OMRÖSTNING","RADERA [TOTAL]",rsDB("omfID"))
              
              Con.ExeCute("DELETE FROM cms_Omrostning_Val WHERE omvFraga = " & CLng(oID))
              Con.ExeCute("DELETE FROM cms_Omrostning_Svar WHERE omsFraga = " & CLng(oID))
              
              rsDB.Delete
            End If
            
            rsDB.Close
          Next
        End if
      
      Set rsDB = Nothing
    Con_Close
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade omröstningarna som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/CMS/Omrostningar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>