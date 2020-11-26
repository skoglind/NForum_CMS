<!--#INCLUDE FILE="cms_MD5.asp"-->

<%
  Dim con_Status
  Dim con
  Dim sSendAlfa
  Dim sAlfa

  Function Con_Open()
    If Not con_Status Then
      Set con = Server.CreateObject("ADODB.Connection")
      con.Mode = 3
      con.Open CONNECTION_STRING
      
      con_Status = True
    End If
  End Function
  
  Function Con_Close()
    If con_Status Then
      con.Close
      Set con = Nothing
      
      con_Status = False
    End If
  End Function

  Function SlumpText(lLength)
    sVals = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,1,2,3,4,5,6,7,8,9,0,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,=,&,%,#,§,?,!,@"
    sArrs = Split(sVals, ",")
    
    Randomize
    For zy = 1 To lLength
      lSlumpKey = CLng(Rnd*UBound(sArrs))
      nStr = nStr & sArrs(lSlumpKey)
    Next
    
    SlumpText = nStr
  End Function
  
  Function ITrim(sText)
    ITrim = sText
  End Function
  
  Function sEncode(sText)
    sEncode = Trim(Server.HTMLEncode(sText & " "))
  End Function
  
  Function jEncode(sText)
    sText = Trim(sText & " ")
  
    sText = Replace(sText, vbCrlf, "\n")
    sText = Replace(sText, "'", "\'")
    sText = Replace(sText, Chr(34), "\" + Chr(34))
    sText = Replace(sText, ",", "\,")
    
    jEncode = sText
  End Function
  
  Function noFnutt(sText)
    nText = sText
    
    nText = Replace(nText, Chr(34), "")   ' - CHAR ( " )
    nText = Replace(nText, Chr(39), "")   ' - CHAR ( ' )
    nText = Replace(nText, Chr(43), "")   ' - CHAR ( + )
    nText = Replace(nText, Chr(45), "")   ' - CHAR ( - )
    nText = Replace(nText, Chr(47), "")   ' - CHAR ( / )
    nText = Replace(nText, Chr(92), "")   ' - CHAR ( \ )
    
    noFnutt = sEncode(nText)
  End Function
  
  Function Make_LC()
    AllaTkn = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z 1 2 3 4 5 6 7 8 9 0"
    Urval = Split(AllaTkn, " ")
    
    Randomize
    For zz = 1 To 25
      nString = nString & Urval(CLng(Rnd*UBound(Urval)))
    Next
  
    Make_LC = "{" & nString & "}"
  End Function
  
  Function GetImgIDByExclusive()
    Set rsEX = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Bild WHERE bUppladdadAv = " & CLng(cCMS_ID) & " AND bSparad = 0 AND bExclusiveID = '" & cCMS_EXKEY & "'"
    rsEX.Open SQL, Con
    
      If rsEX.EOF Then
        lID = 0
      Else
        lID = rsEX("bID")
      End IF
    
    rsEX.Close
    Set rsEX = Nothing
    
    GetImgIDByExclusive = CLng(lID)
  End Function
  
  Function Make_ExclusiveKey()
    Randomize
    sExclusiveString = MD5(CStr(CLng(Rnd*99999) & CLng(Date) & DatePart("h",Now) & DatePart("m",Now) & DatePart("s",Now) & CLng(Rnd*99999) & cCMS_ID))
    Session.Value("CMS_EXKEY") = sExclusiveString
    Make_ExclusiveKey = sExclusiveString
  End Function
  
  Function GetAccess(sAccessFor)
    GetAccess = False
  End Function
  
  Function RoundUp(lNo1, lNo2)
    lSum = CDbl(CDbl(lNo1) / CDbl(lNo2))
    If Round(lSum) < lSum Then lSum = Round(lSum) + 1
    
    RoundUp = lSum
  End Function
  
  Function CutWord(sWord, lLength)
    sWord = Trim(sWord & " ")
    If Len(sWord) > lLength Then sWord = Left(sWord, lLength-3) & "..."
    
    CutWord = sEncode(sWord)
  End Function
  
  Function AddLogg(sObjekt, sAction, lPostID)
    con.Execute("INSERT INTO cms_Logg (lObjekt,lDatum,lAnvID,lAction,lPostID) VALUES('" & sObjekt & "','" & Now & "'," & CLng(cCMS_ID) & ",'" & sAction & "'," & CLng(lPostID) & ")")
  End Function
  
  Function FormToArray(sForm)
    sForm = LCase(sForm)
  
    For Each nObj In Request.Form
      If LCase(Left(nObj,Len(sForm))) = sForm Then nVal = nVal & Trim(Request.Form(nObj)) & " ,"
    Next
    
    If Len(nVal) > 0 Then nVal = Left(nVal, Len(nVal)-1)
    
    FormToArray = Split(nVal,",")
  End Function
  
  Function FormIDToArray(sForm)
    sForm = LCase(sForm)
  
    For Each nObj In Request.Form
      If LCase(Left(nObj,Len(sForm))) = sForm Then nVal = nVal & Right(nObj, Len(nObj)-Len(sForm)) & " ,"
    Next
    
    If Len(nVal) > 0 Then nVal = Left(nVal, Len(nVal)-1)
    
    FormIDToArray = Split(nVal,",")
  End Function
  
  Function IsInArray(aArray, sFind)
    rVal = -1
  
    For zz = 0 To UBound(aArray)
      If CStr(LCase(Trim(aArray(zz)))) = CStr(LCase(Trim(sFind))) Then
        rVal = zz
        Exit For
      End IF
    Next
    
    IsInArray = rVal
  End Function
  
  Function GetFormRequest(sStart, sValue)
    For Each sCHK In Request.Form
      sCHK = Trim(LCase(sCHK))
      If Left(sCHK, 6) = LCase(sStart) Then
        If UCase(Request.Form(sCHK)) = UCase(sValue) Then
          lID = CLng(Right(sCHK, Len(sCHK)-6))
          nID = nID & lID & ","
        End If
      End If
    Next
    If Len(nID) > 0 Then nID = Left(nID, Len(nID)-1)
    
    GetFormRequest = nID
  End Function
  
  Function GetAcc(sDemand)
    bAccess = False
  
    If cCMS_RATTER <> Empty Then
      If InStr(1, cCMS_RATTER, sDemand, vbTextCompare) > 0 Then bAccess = True
    End If
    
    GetAcc = bAccess
  End Function
  
  Function HasAcc(sRatter,sDemand)
    bAccess = False
  
    If sRatter <> Empty Then
      If InStr(1, sRatter, sDemand, vbTextCompare) > 0 Then bAccess = True
    End If
    
    HasAcc = bAccess
  End Function
  
  Function GetAlfa(gAlfa)
    sSendAlfa = Trim(gAlfa)
    
    If UCase(sSendAlfa) = "ARING" Then sSendAlfa = "Å"
    If UCase(sSendAlfa) = "AUML" Then sSendAlfa = "Ä"
    If UCase(sSendAlfa) = "OUML" Then sSendAlfa = "Ö"
    If UCase(sSendAlfa) = "GRIND" Then sSendAlfa = "#"
    If Len(sSendAlfa) > 0 Then UCase(Left(sSendAlfa, 1))
    
    If sSendAlfa <> Empty Then
      Select Case Asc(sSendAlfa)
        Case 35,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,196,197,214
          sAlfa = sSendAlfa
          Select Case UCase(sAlfa)
            Case "Å" : sSendAlfa = "aring"
            Case "Ä" : sSendAlfa = "auml"
            Case "Ö" : sSendAlfa = "ouml"
            Case "#" 
              sSendAlfa = "grind"
              sAlfa = "[0-9_^#.,:;-]_"
          End Select
      End Select
    End IF
  End Function
  
  Function mForm(sObject, sType, lLen)
    mForm = FixGet(Request.Form(sObject), sType, lLen)
  End Function
  
  Function mGet(sObject, sType, lLen)
    mGet = FixGet(Request.QueryString(sObject), sType, lLen)
  End Function
  
  Function FixGet(sValue, sType, lLen)
    Dim sT
    sT = Trim(sValue)
  
    Select Case Trim(UCase(sType))
      Case "123"
        If Not IsNumeric(sT) Or sT = Empty Then sT = 0
        sT = CLng(sT)
      Case "ABC"
        If Not sT = Empty And lLen > 0 And Len(sT) > lLen Then
          sT = Left(sT, lLen)
        End If
      Case "DAT"
        If Not IsDate(sT) Then sT = #2026-12-01#
        sT = CDate(sT)
      Case "CHK"
        If sT = "YES" Then sT = True Else sT = False
    End Select
  
    FixGet = sT
  End Function
  
  Function ListCMSUsers()
    Set rsUsr = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM fsBB_Anv WHERE aS_CMS = 1 ORDER BY aNamn"
    rsUsr.Open SQL, Con
    
      Do Until rsUsr.EOF
        nVals = nVals & "<option value='" & rsUsr("aID") & "' class='levelin'> " & sEncode(rsUsr("aNamn")) & " </option>" & vbCrlf
        rsUsr.MoveNext
      Loop
    
    rsUsr.Close
    Set rsUsr = Nothing
    
    Response.Write nVals
  End Function
  
  Function ListRegion(lSel)
    Set rsReg = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM cms_Region"
    rsReg.Open SQL, Con
    
      Do Until rsReg.EOF
        If rsReg("rHighlight") Then sHL = " style='font-weight: bold;'" Else sHL = ""

        nVals = nVals & "<option value=" & rsReg("rID") & " " & sHL & "> " & sEncode(rsReg("rNamn")) & " </option>" & vbCrlf
        
        rsReg.MoveNext
      Loop
    
    rsReg.Close
    Set rsReg = Nothing
    
    Response.Write nVals
  End FUnction
  
  Function MakeLegal(ByVal sText)
    For t = 1 To Len(sText)
      tkn = Mid(sText, t, 1)
      Select Case Asc(tkn)
        Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90                        ' A-Z
          nText = nText & tkn
        Case 97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122 ' a-z
          nText = nText & tkn
        Case 48,49,50,51,52,53,54,55,56,57                                                                        ' 0-9
          nText = nText & tkn
        Case 132,134,148,142,143,153                                                                              ' åäöÅÄÖ
          nText = nText & tkn
        Case 229,228,246,197,196,214                                                                              ' åäöÅÄÖ (Special??)
          nText = nText & tkn
        Case 95,45,35,94                                                                                          ' _ - # ^
          nText = nText & tkn
      End Select
    Next
    
    MakeLegal = nText
  End Function
  
  Function MakeLegal_Large(ByVal sText)
    For t = 1 To Len(sText)
      tkn = Mid(sText, t, 1)
      Select Case Asc(tkn)
        Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90                        ' A-Z
          nText = nText & tkn
        Case 97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122 ' a-z
          nText = nText & tkn
        Case 48,49,50,51,52,53,54,55,56,57                                                                        ' 0-9
          nText = nText & tkn
        Case 229,228,246,197,196,214                                                                              ' åäöÅÄÖ
          nText = nText & tkn
        Case 95,45,35,94,32,38,44,46,58                                                                            ' _ - # ^ (space) & , . :
          nText = nText & tkn
      End Select
    Next
    
    MakeLegal_Large = nText
  End Function
  
  Function DoUserExist(ByVal sAnvNamn)
    bStatus = True

    nUsers = Con.Execute("SELECT COUNT(*) FROM fsBB_Anv WHERE aAnvNamn = '" & MakeLegal(sAnvNamn) & "'")(0)
    If nUsers > 0 Then bStatus = False
    
    DoUserExist = bStatus
  End Function
  
  Function ImgRemove(myID)
    If Not IsNumeric(myID) Then myID = 0
    myID = CLng(myID)
  
    Set fso = Server.CreateObject("Scripting.FilesystemObject")
    Set rsBB = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Bild WHERE bID = " & CLng(myID)
    rsBB.Open SQL, Con, 1, 3
    
     If Not rsBB.EOF Then
       sFile = UPLOAD_FOLDER & "img_" & Right("0000000000" & rsBB("bID"), 10) & "_original." & rsBB("bTyp")
       If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
       
       For zx = 1 To lstImgSize(0)
         sFile = UPLOAD_FOLDER & Replace(Replace(lstImgSize(zx), ",", "x"), "LOGIN_", "") & "/img_" & Right("0000000000" & rsBB("bID"), 10) & ".png"
         If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
       Next
     End If 
    
    rsBB.Close
    Set rsBB = Nothing
    Set fso = Nothing
    
    Con.ExeCute("DELETE FROM cms_Bild WHERE bID = " & CLng(myID))
  End Function
  
  Function RemoveAllFile(sID, sAndelse)
    sStartOfFile = "img_" & Right("0000000000" & sID, 10)
    RemoveFile(sStartOfFile & "_original." & sAndelse)
  End Function
  
  Function ImgDoRenew(myID, sSize)
    sSizes = Con.ExeCute("SELECT bInSizes FROM cms_Bild WHERE bID = " & CLng(myID))(0)
    
    If InStr(sSizes, sSize) > 0 Then
    Else
      mSize = Split(sSize, ",")
      ImgResize myID, mSize(0), mSize(1), 80
      
      sSizes = sSizes & ";" & sSize
      Con.ExeCute("UPDATE cms_Bild Set bInSizes = '" & sSizes & "' WHERE bID = " & CLng(myID))
    End If
  End Function
  
  Function ImgOriginal(lID)
    Set rsBB = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Bild WHERE bID = " & CLng(lID)
    rsBB.Open SQL, Con
    
      If rsBB.EOF Then
        ImgOriginal = "NO_IMG"
      Else
        ImgOriginal = "img_" & Right("0000000000" & lID, 10) & "_original." & rsBB("bTyp")
      End If
    
    rsBB.Close
    Set rsBB = Nothing
  End Function
  
  Function ImgResize(sImgID, lWidth, lHeight, lCompression)
    sImage = UPLOAD_FOLDER & ImgOriginal(sImgID)
  
    Set Jpeg = Server.CreateObject("Persits.Jpeg")
      Jpeg.Open CStr(sImage)
      Jpeg.Canvas.Brush.Color = &HFFFFFF
      Jpeg.Interpolation = 1
      Jpeg.Quality = lCompression
      Jpeg.Progressive = True
      Jpeg.PNGOutput = True
      
      oWidth  = Jpeg.OriginalWidth
      oHeight = Jpeg.OriginalHeight
    
      Jpeg.PreserveAspectRatio = True
      
      nWidth_Diff = oWidth - lWidth
      nHeight_Diff = oHeight - lHeight
      
      If nWidth_Diff > nHeight_Diff Then
        If nWidth_Diff > -1 Then Jpeg.Width = lWidth
      Else
        If nHeight_Diff > -1 Then Jpeg.Height = lHeight
      End If
      
      nWidth  = Jpeg.Width
      nHeight = Jpeg.Height
      
      If nWidth < lWidth Then
        nx0 = -((lWidth - nWidth) / 2)
        nx1 = nWidth + ((lWidth - nWidth) / 2)
      Else
        nx0 = 0
        nx1 = nWidth
      End If
      
      If nHeight < lHeight Then
        ny0 = -((lHeight - nHeight) / 2)
        ny1 = nHeight + ((lHeight - nHeight) / 2)
      Else
        ny0 = 0
        ny1 = nHeight
      End If
      
      Jpeg.Crop nx0, ny0, nx1, ny1
      Jpeg.Crop 0, 0, lWidth, lHeight
      
      sFileSave = UPLOAD_FOLDER & lWidth & "x" & lHeight & "\img_" &  Right("0000000000" & sImgID, 10) & ".png"
      
      Jpeg.Save sFileSave
    Set Jpeg = Nothing
    
    ImgResize = sFileSave
  End Function
%>