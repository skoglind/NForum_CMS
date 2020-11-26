<%
Function lstStatus(lNo)
  iBound = 5

  Select Case CLng(lNo)
    Case  0   : sData = "Raderad"
    Case  1   : sData = "Under bearbetning"
    Case  2   : sData = "Inväntar publicering"
    Case  3   : sData = "Publicering nekad"
    Case  4   : sData = "Publicerad"
    Case Else : sData = CLng(iBound)
  End Select

  lstStatus = sData
End Function

Function lstKategori(lNo)
  iBound = 18

  Select Case CLng(lNo)
    Case  1   : sData = "Nintendo allmänt"
    Case  2   : sData = "Bärbart"
    Case  3   : sData = "Stationärt"
    Case  4   : sData = "Gameboy (Original)"
    Case  5   : sData = "Gameboy Color"
    Case  6   : sData = "Gameboy Advance"
    Case  7   : sData = "Nintendo DS"
    Case  8   : sData = "Nintendo (Original)"
    Case  9   : sData = "Super Nintendo"
    Case 10   : sData = "Nintendo 64"
    Case 11   : sData = "Nintendo Gamcube"
    Case 12   : sData = "Nintendo Wii"
    Case 13   : sData = "Virtual Boy"
    Case 14   : sData = "Game & Watch"
    Case 15   : sData = "N-Forum.se/Gameboy.nu"
    Case 16   : sData = "Övrigt"
    Case 17   : sData = "Color TV-Game"
    Case 18   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKategori = sData
End Function

Function lstKonsol(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "Nintendo Gameboy"
    Case  2   : sData = "Nintendo Gameboy Color"
    Case  3   : sData = "Nintendo Gameboy Advance"
    Case  4   : sData = "Nintendo DS"
    Case  5   : sData = "Nintendo (8-bit)"
    Case  6   : sData = "Super Nintendo"
    Case  7   : sData = "Nintendo 64"
    Case  8   : sData = "Nintendo Gamecube"
    Case  9   : sData = "Nintendo Wii"
    Case 10   : sData = "Nintendo Virtual Boy"
    Case 11   : sData = "Nintendo Game & Watch"
    Case 12   : sData = "Color TV-Game"
    Case 13   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsol = sData
End Function

Function lstKonsolShort(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "Gameboy"
    Case  2   : sData = "Gameboy Color"
    Case  3   : sData = "Gameboy Advance"
    Case  4   : sData = "Nintendo DS"
    Case  5   : sData = "Nintendo (8-bit)"
    Case  6   : sData = "Super Nintendo"
    Case  7   : sData = "Nintendo 64"
    Case  8   : sData = "Gamecube"
    Case  9   : sData = "Nintendo Wii"
    Case 10   : sData = "Virtual Boy"
    Case 11   : sData = "Game & Watch"
    Case 12   : sData = "Color TV-Game"
    Case 13   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsolShort = sData
End Function

Function lstKonsolXShort(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "GB"
    Case  2   : sData = "GBC"
    Case  3   : sData = "GBA"
    Case  4   : sData = "NDS"
    Case  5   : sData = "NES"
    Case  6   : sData = "SNES"
    Case  7   : sData = "N64"
    Case  8   : sData = "NGC"
    Case  9   : sData = "WII"
    Case 10   : sData = "VB"
    Case 11   : sData = "G&W"
    Case 12   : sData = "CTVG"
    Case 13   : sData = "3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsolXShort = sData
End Function

Function lstRegion(lNo)
  iBound = 11

  Select Case CLng(lNo)
    Case  1   : sData = "Europa"
    Case  2   : sData = "USA"
    Case  3   : sData = "Japan"
    Case  4   : sData = "Skandinavien"
    Case  5   : sData = "Australien"
    Case  6   : sData = "Sverige"
    Case  7   : sData = "Tyskland"
    Case  8   : sData = "Kanada"
    Case  9   : sData = "Frankrike"
    Case 10   : sData = "Spanien"
    Case 11   : sData = "Storbritannien"
    Case Else : sData = CLng(iBound)
  End Select

  lstRegion = sData
End Function

Function lstImgSize(lNo)
  iBound = 11

  Select Case CLng(lNo)
    Case  1   : sData = "28,28"
    Case  2   : sData = "50,45"
    Case  3   : sData = "50,50"
    Case  4   : sData = "80,80"
    Case  5   : sData = "100,100"
    Case  6   : sData = "150,150"
    Case  7   : sData = "320,240"
    Case  8   : sData = "LOGIN_640,480"
    Case  9   : sData = "LOGIN_800,600"
    Case 10   : sData = "LOGIN_1024,768"
    Case 11   : sData = "LOGIN_1280,1024"
    Case 12   : sData = "200,150"
    Case Else : sData = CLng(iBound)
  End Select

  lstImgSize = sData
End Function
%>