<% measureStart = Timer %>

<!--#INCLUDE FILE="cms_Config.asp"-->
<!--#INCLUDE FILE="cms_Constant.asp"-->
<!--#INCLUDE FILE="cms_Lists.asp"-->
<!--#INCLUDE FILE="cms_Functions.asp"-->

<%
Response.Buffer = True
Response.Expires = 0

lID = mGet("e","123",0)
lW  = mGet("w","123",0)
lH  = mGet("h","123",0)
sF  = mGet("filter","ABC",25)
sEr = LCase(mGet("err","ABC",2))

  Con_Open

  sErrorFile = Server.MapPath("/design/noimg.png")
  
    sOriginal = ImgOriginal(lID)
      
    If sOriginal = "NO_IMG" Then
      cSTREAMFILE = sErrorFile
    Else
      bShowThumb = True
      For zx =  1 To lstImgSize(0)
        sCompVal = lstImgSize(zx)
        If Session.Value("CMS_LOGIN") Then sCompVal = Replace(sCompVal, "LOGIN_", "")
        
        If sCompVal = CStr(lW & "," & lH) Then bShowThumb = False
      Next
  
      sFullOrg  = UPLOAD_FOLDER & sOriginal
      
      If bShowThumb Then
        sFile = "150x150\img_" & Right("0000000000" & lID, 10) & ".png"
        lW = 150
        lH = 150
      Else
        sFile = lW & "x" & lH & "\img_" & Right("0000000000" & lID, 10) & ".png"
      End If
      
      sFullFile = UPLOAD_FOLDER & sFile
      
      Set fso = Server.CreateObject("Scripting.FileSystemObject")
        bHasOriginal  = True
        If Not fso.FileExists(sFullOrg) Then bHasOriginal = False
      Set fso = Nothing
      
      If bHasOriginal Then 
        ImgDoRenew lID, lW & "," & lH
        
        cSTREAMFILE = sFullFile
      Else
        cSTREAMFILE = sErrorFile
      End If
    End If
    
  Con_Close
  
  Response.Clear
  
  If sEr = "no" Then If cSTREAMFILE = sErrorFile Then Response.End

  'Const adTypeBinary = 1 
  'cCONTENTTYPE = "image/jpeg"
  '
  'Response.Contenttype = cCONTENTTYPE
  '  
  '  Set Stream = server.CreateObject("ADODB.Stream") 
  '  Stream.Type = adTypeBinary 
  '  Stream.Open 
  '  Stream.LoadFromFile cSTREAMFILE
  '  While Not Stream.EOS 
  '    Response.BinaryWrite Stream.Read(1024 * 64) 
  '  Wend 
  '  Stream.Close 
  '  Set Stream = Nothing 
    
  'Response.Flush 
  'Response.End
  
  Set Jpeg = Server.CreateObject("Persits.Jpeg")
    Jpeg.Open(cSTREAMFILE)
    Jpeg.PNGOutput = True
    
    'measureStop = Timer
    'measure = FormatNumber(measureStop - measureStart, 3)
    
    'Select Case Trim(LCase(sF))
    '  Case "grayscale"
    '    Jpeg.Grayscale 1
    '  Case "timer"
    '    Jpeg.Canvas.Font.Color  = &H000000
    '    Jpeg.Canvas.Font.Size   = 12
    '    Jpeg.Canvas.PrintTextEx measure & " seconds", 2, 12, "c:\Windows\Fonts\Arial.ttf"
    '  Case "sharpen"
    '    Jpeg.Interpolation = 2
    'End Select
    
    'If lW > 600 Then
    '  Jpeg.Canvas.DrawPNG Jpeg.Width - 340, Jpeg.Height - 115, Server.MapPath("../bilder") & "\stamp_large.png"
    'ElseIf lW > 150 Then
    '  Jpeg.Canvas.DrawPNG Jpeg.Width - 140, Jpeg.Height - 44, Server.MapPath("../bilder") & "\stamp.png"
    'End If
    
  Jpeg.SendBinary
  Set Jpeg = Nothing
%>