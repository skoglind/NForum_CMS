<%
  ' #### MILJÖVARIABLER ####
    CMS_SITENAME          = "N-Forum"
    CMS_SITEADDR          = "http://www.n-forum.se"
    CMS_HALT              = False
    CONNECTION_STRING     = "Provider=SQLOLEDB;Data Source=creeper\SQLExpress2008;Initial Catalog=db_NForum;User Id=*****;Password=*****"

  ' #### INLOGGNING ####
    LOGIN_TABLE           = "fsBB_Anv"
    LOGIN_USR             = "aAnvNamn"
    LOGIN_PWS             = "aPassWd"
    LOGIN_BEHORIGHET      = "aCMS"
    config_Hash_Salt_1    = "***"
    config_Hash_Salt_2    = "***"
    
    LOGIN_FORSOK          = 5
    LOGIN_INOM            = 30
    LOGIN_MINUTER         = 60
    
  ' #### FSBB FORUM ####  
    FSBB_FORUM             = True
    FSBB_FORUMADDR         = "http://www.n-forum.se.nu/avdelning/forum"
    FSBB_DEFAULTUSER       = 8
    FSBB_NEWSFORUM         = 13

  ' #### PUBLICERING ####
    PUBL_TID               = #10:00:00#
    
  ' #### BILDUPPLADDNING ####
    UPLOAD_MAXSIZE         = 3145728 ' (3MB)
    UPLOAD_FOLDER          = "C:\WebbRoot\N-Forum.se\bilder\" 
    UPLOAD_FAKE            = "../Bilder/"
    
  ' #### PAGING ####
    MAXPERPAGE             = 20
%>