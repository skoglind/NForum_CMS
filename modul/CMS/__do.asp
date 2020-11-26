<!--#INCLUDE FILE="../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../cms_Lists.asp"-->
<!--#INCLUDE FILE="__do_Func.asp"-->

<%
If Not GetAcc("CMS0") Then Response.Redirect("/")
%>

<%
sAction       = Trim(LCase(Request.QueryString("a")))
sExtraAction  = Trim(LCase(Request.QueryString("ea")))

' #### REMEMBER ####
sRebuild = "x="
' ##################

Select Case sAction
  Case "save" ' Spara
    lID         = mForm("vID", "123", 0)
    sTitel      = mForm("vTitel", "ABC", 30)
    sText       = mForm("vText", "ABC", 0)
    lSortNr     = mForm("vSortNr", "123", 0)
    
    Con_Open
      Set rsDB = Server.CreateObject("ADODB.RecordSet")
      SQL = "SELECT * FROM cms_InfoBlock WHERE ifID = " & CLng(lID)
      rsDB.Open SQL, Con, 1, 3
    
      ' #### FELHANTERING ####
        bErr = False
      
        If Len(sTitel) < 1 Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon titel.</p>"
        If Len(sText) < 1  Then bErr = True : nMessage = "<p>Inget har lagrats i databasen då du inte har angett någon text.</p>"
        
        If Not rsDB.EOF Then
          bIsNew = False
        Else
          bIsNew = True
        End If
        
        If bErr Then
          Response.Write("<script type='text/javascript'>parent.savefailed('" & nMessage & "');</script>")
          Response.Write("<script type='text/javascript'>location.href='../../_awaiting.asp';</script>")
          Response.End
        End If
      ' ######################
      
        If rsDB.EOF Then
          rsDB.AddNew
        End If
        
        rsDB("ifTitel")       = sTitel
        rsDB("ifTextM")       = sText
        
        rsDB("ifSortNr")      = lSortNr
        
        rsDB("ifDatumSparad") = Now
        saveDate = "Sparad (" & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & ")"
        
        rsDB.Update
      
        lID = rsDB("ifID")
        If bIsNew Then AddLogg "INFOBLOCK","SKAPA",lID
        
      rsDB.Close
      Set rsDB = Nothing
    Con_Close
    
    Select Case sExtraAction
      Case "continue"
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",true,'" & "_edit.asp');</script>")
      Case "return"
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",true,'" & "default.asp');</script>")
      Case Else
        Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",false,'');</script>")
    End Select
    
    Response.Write("<script type='text/javascript'>location.href='../../_awaiting.asp';</script>")
  Case "del" ' Radera
    If GetAcc("CMS0") Then
      Con_Open
        lID = mGet("e", "123", 0)
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM cms_InfoBlock WHERE ifID = " & CLng(lID)
        rsDB.Open SQL, Con, 1, 3
      
          If Not rsDB.EOF Then
            rsDB.Delete
          End If
        
        rsDB.Close
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Response.Redirect("default.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../_awaiting.asp';</script>")
End Select
%>