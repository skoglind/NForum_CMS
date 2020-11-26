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
    ' Njet
  Case "del" ' Radera
    If GetAcc("CMS3") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM fsBB_Anmal WHERE anID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
              rsDB.Delete
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Radering slutförd</h2><p>De markerade anmälningarna som du hade behörighet att radera är nu borta.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/fsBB/Anmalningar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case "notera" ' Behandla
    If GetAcc("CMS3") Then
      Con_Open
        allID = Split(GetFormRequest("chk_id", "YES"), ",")
        Set rsDB = Server.CreateObject("ADODB.RecordSet")
        
          For Each oID IN allID
            SQL = "SELECT * FROM fsBB_Anmal WHERE anID = " & CLng(oID)
            rsDB.Open SQL, Con, 1, 3
              rsDB("anNoterad") = True
              rsDB.Update
            rsDB.Close
          Next
        
        Set rsDB = Nothing
      Con_Close
    End If
    
    Session.value("PBM_Message")    = "<h2>Information: Åtgärden slutförd</h2><p>De markerade anmälningarna har nu markerats som behandlade.</p><p>Klicka på ""fortsätt"" för att gå vidare...</p>"
    Session.value("PBM_Lank")       = "modul/fsBB/Anmalningar/_show.asp?" & sRebuild
  
    Response.Redirect("../../../_message.asp")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>