<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS333") Then Response.Redirect("/")
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
    sName       = mForm("vNamn", "ABC", 30)
    sInfo       = mForm("vInfo", "ABC", 255)
    
    lSortering  = mForm("vSortering", "123", 0)
    lSortNr     = mForm("vSortNr", "123", 0)
    
    bNoAllView     = mForm("vNoAllView", "CHK", 0)
    bSplitterBefore= mForm("vSplitterBefore", "CHK", 0)
    bHideMe        = mForm("vHideMe", "CHK", 0)
    bGroup         = mForm("vGroup", "CHK", 0)
    
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM fsBB_Forum WHERE fID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sName) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett något namn för forum.</p>"
        
        If Not rsDB.EOF Then
          lPost = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tForum = " & CLng(lID))(0)
          If lPost > 0 And bGroup Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte kan göra ett forum med poster till en grupp.</p>"
        
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
        
        rsDB("fName")        = sName
        rsDB("fInfo")        = sInfo
        
        rsDB("fSortering")   = lSortering
        rsDB("fSortNr")      = lSortNr
        
        rsDB("fNoAllView")     = bNoAllView
        rsDB("fSplitterBefore")= bSplitterBefore
        rsDB("fHideMe")        = bHideMe
        rsDB("fGroup")         = bGroup
        
        If Not bGroup Then
          ' #### NY TRÅD ####
            sNewThread = ";"
            For Each lNewThread In Request.Form("newthread")
              sNewThread = sNewThread & lNewThread & ";"
            Next
            If Len(sNewThread) < 2 Then sNewThread = "0"
          ' ###############
          
          ' #### NYTT INLÄGG ####
            sNewReply = ";"
            For Each lNewReply In Request.Form("newreply")
              sNewReply = sNewReply & lNewReply & ";"
            Next
            If Len(sNewReply) < 2 Then sNewReply = "0"
          ' ###############
          
          ' #### VISA ####
            sView = ";"
            For Each lView In Request.Form("view")
              sView = sView & lView & ";"
            Next
            If Len(sView) < 2 Then sView = "0"
          ' ###############
          
          ' #### MODERATOR ####
            sMod = ";"
            For Each lMod In Request.Form("mod")
              sMod = sMod & lMod & ";"
            Next
            If Len(sMod) < 2 Then sMod = "0"
          ' ###############
          
          rsDB("fSec_NewThread")  = sNewThread
          rsDB("fSec_NewReply")   = sNewReply
          rsDB("fSec_View")       = sView
          rsDB("fSec_Mod")        = sMod
          rsDB("fGroupForums")    = "0"
        Else
          ' #### GRUPPER ####
            For Each lGroup In Request.Form("group")
              sGroup = sGroup & lGroup & ","
            Next
            If Len(sGroup) > 0 Then sGroup = Left(sGroup, Len(sGroup)-1)
            If sGroup = "" Then sGroup = "0"
          ' ###############
          
          rsDB("fSec_View")       = "0"
          rsDB("fGroupForums")    = sGroup
        End If
        
        rsDB("fDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("fID")
        If bIsNew Then AddLogg "FORUM","SKAPA",lID
        
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
    If GetAcc("CMS333") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM fsBB_Forum WHERE fGroup = 0 AND fID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
          
            If Not rsDB.Eof Then
              lPost = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tForum = " & CLng(oID))(0)
              If lPost = 0 Then rsDB.Delete
            End If
            
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade forumkategorierna som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/fsBB/Forum/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>