<%
Function IsBehorig(sModul, lID)
  Dim bStatus

  bStatus = True
  
  Set rsCHK = Server.CreateObject("ADODB.RecordSet")
   
    Select Case Trim(UCase(sModul))
    Case "NEWS"
      SQL = "SELECT * FROM cms_Nyheter WHERE nID = " & CLng(lID)
      bAlla = GetAcc("CMS111")
      bPubl = GetAcc("CMS11")
      
      rsCHK.Open SQL, Con
   
        If Not rsCHK.EOF Then
          If Not bAlla And CLng(rsCHK("nSkapadAv")) <> CLng(cCMS_ID) Then bStatus = False
          If Not bPubl And (rsCHK("nStatus") = 2 Or rsCHK("nStatus") = 4) Then bStatus = False
        End If
    
      rsCHK.Close
    Case "REC"
      SQL = "SELECT * FROM cms_Recensioner WHERE rID = " & CLng(lID)
      bAlla = GetAcc("CMS111")
      bPubl = GetAcc("CMS11")
      
      rsCHK.Open SQL, Con
   
        If Not rsCHK.EOF Then
          If Not bAlla And CLng(rsCHK("rSkapadAv")) <> CLng(cCMS_ID) Then bStatus = False
          If Not bPubl And (rsCHK("rStatus") = 2 Or rsCHK("rStatus") = 4) Then bStatus = False
        End If
    
      rsCHK.Close
    Case "ART"
      SQL = "SELECT * FROM cms_Artiklar WHERE aaID = " & CLng(lID)
      bAlla = GetAcc("CMS111")
      bPubl = GetAcc("CMS11")
      
      rsCHK.Open SQL, Con
   
        If Not rsCHK.EOF Then
          If Not bAlla And CLng(rsCHK("aaSkapadAv")) <> CLng(cCMS_ID) Then bStatus = False
          If Not bPubl And (rsCHK("aaStatus") = 2 Or rsCHK("aaStatus") = 4) Then bStatus = False
        End If
    
      rsCHK.Close
    Case "GAME"
      bStatus = True
    Case Else
      bStatus = False
    End SElect
  
  Set rsCHK = Nothing
  
  IsBehorig = bStatus
End Function
%>