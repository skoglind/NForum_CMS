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

' #### REMEMBER ####

' ##################

Select Case sAction
  Case "save" ' Spara
    sTextM     = mForm("vText", "ABC", 0)
    
    ' #### FELHANTERING ####
      bErr = False
    
      If Len(sTextM) < 1 Then bErr = True : nMessage = "<p>Listan är tom!</p>"
      
      If bErr Then
        Response.Write("<script type='text/javascript'>parent.savefailed('" & nMessage & "');</script>")
        Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
        Response.End
      End If
    ' ######################
    
    
    ' #### LÄS UT LISTAN ####
    
      Function cntArray(element, arr)
        Dim iAnt
        
        For zx = 0 To Ubound(arr) 
          If LCase(Trim(arr(zx))) = LCase(Trim(element)) Then iAnt = iAnt + 1
        Next
        
        cntArray = iAnt
      End Function
      
      Function dbHit(element)
        dbHit = False
      End Function
    
      aSpelTitlel = Split(sTextM ,vbCrlf)
      
      For zz = 0 To Ubound(aSpelTitlel)
        If cntArray(aSpelTitlel(zz),aSpelTitlel) > 1 Then 
          aSpelTitlel(zz) = ""
        Else
          If Len(aSpelTitlel(zz)) > 3 Then
            If dbHit(aSpelTitlel(zz)) Then
              aSpelTitlel(zz) = ""
            Else
              ' #### ANVÄND ####
            End If
          End if
        End If
      Next
    
    ' #######################
    
    saveDate  = Now
    lID       = 0
    
    Response.Write("<script type='text/javascript'>parent.savefinished('" & saveDate & "'," & lID & ",false,'');</script>")
    'Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
  Case Else
    Response.Write("<script type='text/javascript'>location.href='../../../_awaiting.asp';</script>")
End Select
%>