<%
Response.addHeader "pragma","no-cache"
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.CacheControl = "no-cache"
%>

<!--#INCLUDE FILE="../../../cms_Config.asp"-->
<!--#INCLUDE FILE="../../../cms_Constant.asp"-->
<!--#INCLUDE FILE="../../../cms_Functions.asp"-->
<!--#INCLUDE FILE="../../../cms_Lists.asp"-->

<%
If Not GetAcc("CMS4") Then Response.Redirect("/")
%>

<%
cADD_ID = mGet("e","123",0)
Con_Open
  Set rsDB = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM cms_Foretag LEFT JOIN cms_Bild ON cms_Foretag.fLogga = cms_Bild.bID WHERE fID = " & CLng(cADD_ID)
    rsDB.Open SQL, Con
    
      anyHit = True
      If rsDB.EOF Then
        anyHit = False
        cADD_ID = 0
      Else
        cADD_ID = rsDB("fID")
        If IsNull(rsDB("bID")) Then 
          anyHit = False
        Else
          bildID  = CLng(rsDB("bID"))
          bildTyp = Trim(rsDB("bTyp"))
        End If
      End IF
    
    rsDB.Close
  Set rsDB = Nothing
  
  If CLng(cADD_ID) = 0 Then
    bildID = GetImgIDByExclusive()
    If CLng(bildID) > 0 Then anyHit = True
  End If
  
Con_Close
%>


<form id="smallupload" method="POST" enctype="multipart/form-data">
  <div class="field"><input type="file" name="thaFile" size="16"></div>
  <div class="field"><input type="button" value="Ladda upp" onclick="smalluploadimg('smallupload');"></div>
  <input type="hidden" value=<% = cADD_ID %> name="lID">
  <input type="hidden" value="foretag" name="sArea">
</form>

<% If anyHit Then %>
  <%
  Con_Open
    sOriginal = ImgOriginal(bildID)
    If sOriginal = "NO_IMG" Then
      anyPic = False
    Else
      anyPic = True
    End IF
  Con_Close
  
  sFilNamn = "/cms_Img.asp?e=" & bildID & "&w=150&h=150"
  %>
  
  <div class="innerseparator"> </div>
  <img src="<% If anyPic Then Response.Write(sFilNamn) %>" style="float: left; border: solid 1px #CCC; width: 150px; height: 150px; background-color: #FFF; display: block; margin: 0 12px 5px 12px;">
  <div class="field"><input type="button" value="Radera logotype" onclick="if(confirm('Vill du radera logotypen?')){smalldeleteimg(<% = cADD_ID %>,'foretag');}"></div>
<% End If %>