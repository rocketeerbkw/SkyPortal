<!--#INCLUDE file="config.asp" --><% 
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
dim do_app, app_version, app_id 
bUninstall = false
bReinstall = false

':: leave this as is.
strModTablePrefix = ""
app_version = "0.8"
do_app = true
incForumFp = false
%>

<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="includes/inc_DBFunctions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	menu_fp() 
	%></td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
if intIsSuperAdmin then
  if incForumFp then
    createForums()
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the fp_forums.asp ""include"" file<br />")
    Response.Write("to your fp_custom.asp file in order<br />")
    Response.Write("to install this module</p><p>&nbsp;</p>")
    spThemeBlock1_close(intSkin)
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
  end if
else
  Response.Write("<p>&nbsp;</p>")
  spThemeBlock1_open(intSkin)
  Response.Write("<p>&nbsp;</p><p>You must be logged in as a <b>Super Admin</b>")
  Response.Write(" in order to install this module</p><p>&nbsp;</p>")
  spThemeBlock1_close(intSkin)
  Response.Write("<p>&nbsp;</p>")
  Response.Write("<p>&nbsp;</p>")
end if

%>
	</td>
    <td class="rightPgCol">
	<% intSkin = getSkin(intSubSkin,3) %>
	</td>
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" --><%

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::		SUBROUTINES BELOW HERE
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub forum_Upgrades()
  forum_Upgrade07()
end sub

sub createForums()
  spThemeBlock1_open(intSkin)
  response.Write("<hr /><h3>FORUMS MODULE</h3><br />")
  
  'check if app is existing
  sSql = "SELECT APP_NAME,APP_ID,APP_VERSION FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall or bReinstall then
      uninstall_Forums()
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing

 if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "0.8"
	   ':: current version
     case "0.7"
	   updateVersion app_version,"forums"
	 case else
	   forum_Upgrades()
   end select
 elseif not bUninstall then
    addApp()

    crCatTbl()
    crForumTbl()
    crTopicTbl()
    crReplyTbl()
    crPollTbls()
    crReportedTbl()
    crArchiveTbls()
    crModeratorTbl()
	crAllowedMembers()
	
    addFp()
	
	crIndexes()
	'crRelationships()

    Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
    forum_Upgrades()
 end if
 if not bUninstall then
  response.Write("<hr /><h3>Forums Module Installed</h3><br /><br />")
 else
  response.Write("<hr /><h3>Forums Module Uninstalled</h3><br /><br />")
 end if
  response.Write("<b>Be sure to delete this file (createForums.asp) from your server!</b><br /><br />")
  response.Write("<a href=""default.asp""><b>Continue</b></a><br /><br /><br /><br />")
  spThemeBlock1_close(intSkin)
  Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
end sub

sub crIndexes()
':: create indexes :::::::::::::::::::
redim indexes(3)
indexes(0) = "CREATE INDEX [FAR_CAT_ID] ON [" & strTablePrefix & "ARCHIVE_REPLY]([CAT_ID]);"
indexes(1) = "CREATE INDEX [FAR_FORUM_ID] ON [" & strTablePrefix & "ARCHIVE_REPLY]([FORUM_ID]);"
indexes(2) = "CREATE INDEX [FAR_REPLY_ID] ON [" & strTablePrefix & "ARCHIVE_REPLY]([REPLY_ID]);"
indexes(3) = "CREATE INDEX [FAR_TOPIC_ID] ON [" & strTablePrefix & "ARCHIVE_REPLY]([TOPIC_ID]);"
createIndx(indexes)

redim indexes(2)
indexes(0) = "CREATE INDEX [FAT_CAT_ID] ON [" & strTablePrefix & "ARCHIVE_TOPICS]([CAT_ID]);"
indexes(1) = "CREATE INDEX [FAT_FORUM_ID] ON [" & strTablePrefix & "ARCHIVE_TOPICS]([FORUM_ID]);"
indexes(2) = "CREATE INDEX [FAT_TOPIC_ID] ON [" & strTablePrefix & "ARCHIVE_TOPICS]([TOPIC_ID]);"
createIndx(indexes)

createIndex("CREATE INDEX [FORUM_AM] ON [" & strTablePrefix & "ALLOWED_MEMBERS]([FORUM_ID],[MEMBER_ID]);")

redim indexes(0)
indexes(0) = "CREATE INDEX [FORUM_CATEGORYCAT_STATUS] ON [" & strTablePrefix & "CATEGORY]([CAT_STATUS]);"
createIndx(indexes)

redim indexes(3)
indexes(0) = "CREATE INDEX [F_Cat] ON [" & strTablePrefix & "FORUM]([CAT_ID]);"
indexes(1) = "CREATE INDEX [F_LAST_POST] ON [" & strTablePrefix & "FORUM]([F_LAST_POST]);"
indexes(2) = "CREATE INDEX [F_LAST_POST_AUTHOR] ON [" & strTablePrefix & "FORUM]([F_LAST_POST_AUTHOR]);"
indexes(3) = "CREATE INDEX [F_ORDER] ON [" & strTablePrefix & "FORUM]([FORUM_ORDER]);"
createIndx(indexes)

redim indexes(6)
indexes(0) = "CREATE INDEX [T_LAST_POST] ON [" & strTablePrefix & "TOPICS]([T_LAST_POST]);"
indexes(1) = "CREATE INDEX [CAT_ID] ON [" & strTablePrefix & "TOPICS]([CAT_ID]);"
indexes(2) = "CREATE INDEX [Forum_id] ON [" & strTablePrefix & "TOPICS]([FORUM_ID]);"
indexes(3) = "CREATE INDEX [T_AUTHOR] ON [" & strTablePrefix & "TOPICS]([T_AUTHOR]);"
indexes(4) = "CREATE INDEX [T_DATE] ON [" & strTablePrefix & "TOPICS]([T_DATE]);"
indexes(5) = "CREATE INDEX [T_STATUS] ON [" & strTablePrefix & "TOPICS]([T_STATUS]);"
indexes(6) = "CREATE INDEX [T_LAST_POST_AUTHOR] ON [" & strTablePrefix & "TOPICS]([T_LAST_POST_AUTHOR]);"
'indexes(7) = "CREATE INDEX [PrimaryKey] ON [" & strTablePrefix & "TOPICS]([FORUM_ID],[TOPIC_ID],[CAT_ID]) WITH PRIMARY;"
createIndx(indexes)

redim indexes(5)
indexes(0) = "CREATE INDEX [REPLY_ID] ON [" & strTablePrefix & "REPLY]([REPLY_ID]);"
indexes(1) = "CREATE INDEX [CAT_ID] ON [" & strTablePrefix & "REPLY]([CAT_ID]);"
indexes(2) = "CREATE INDEX [FORUM_ID] ON [" & strTablePrefix & "REPLY]([FORUM_ID]);"
indexes(3) = "CREATE INDEX [R_DATE] ON [" & strTablePrefix & "REPLY]([R_DATE]);"
indexes(4) = "CREATE INDEX [R_AUTHOR] ON [" & strTablePrefix & "REPLY]([R_AUTHOR]);"
indexes(5) = "CREATE INDEX [Topic_ID] ON [" & strTablePrefix & "REPLY]([TOPIC_ID]);"
'indexes(6) = "CREATE INDEX [PrimaryKey] ON [" & strTablePrefix & "REPLY]([FORUM_ID],[TOPIC_ID],[REPLY_ID],[CAT_ID]) WITH PRIMARY;"
createIndx(indexes)

createIndex("CREATE INDEX [POLL_ID] ON [" & strTablePrefix & "POLLS]([POLL_ID]);")

redim indexes(1)
indexes(0) = "CREATE INDEX [MEMBER_ID] ON [" & strTablePrefix & "POLL_ANS]([MEMBER_ID]);"
indexes(1) = "CREATE INDEX [POLL_ID] ON [" & strTablePrefix & "POLL_ANS]([POLL_ID]);"
createIndx(indexes)

redim indexes(1)
indexes(0) = "CREATE INDEX [FORUM_ID] ON [" & strTablePrefix & "MODERATOR]([FORUM_ID]);"
indexes(1) = "CREATE INDEX [MEMBER_ID] ON [" & strTablePrefix & "MODERATOR]([MEMBER_ID]);"
createIndx(indexes)
end sub
	
sub addApp()
  'create the app
  response.Write("<hr /><h4>Update PORTAL_APPS</h4><br />")
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_iNAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW],[APP_VERSION]"
  arrData(2) = "'forums','forums',1,0,'1,2,3',1,1,'config_forums','link.asp?topicID=','" & app_version & "'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub addFp()
':::::::::::: add forum items to front page :::::::::::::::::::::::::::
  response.Write("<hr /><h4>Add forum items to front page</h4><br />")
	redim arrData(4)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[fp_name],[fp_iname],[fp_function],[fp_active],[fp_column],[fp_desc],[fp_groups],[APP_ID]"
	arrData(2) = "'Forum recent topics','forum_topics','f_topics_sm',1,4,'Recent topics from the forum.','3'," & app_id & ""
	arrData(3) = "'Featured Polls','forum_polls','f_polls_fp',1,4,'Featured polls from the forums.','3'," & app_id & ""
	arrData(4) = "'Forum News','forum_news','f_news_fp',1,2,'Website News.','3'," & app_id & ""
	populateB(arrData)
	dbHits = dbHits + 1
	
	sSQL = "SELECT fp_leftcol, fp_maincol, fp_rightcol, fp_mainsticky FROM " & strTablePrefix & "FP_USERS WHERE fp_uid=0"
	set fpRS = my_Conn.execute(sSQL)
	dbHits = dbHits + 1
	if not fpRS.eof then
	  leftcol = trim(fpRS("fp_leftcol"))
	  maincol = trim(fpRS("fp_maincol"))
	  rightcol = trim(fpRS("fp_rightcol"))
	  mainsticky = trim(fpRS("fp_mainsticky"))
	  if leftcol <> "" then
	    leftcol = leftcol & ","
	  end if
	  if maincol <> "" then
	    maincol = maincol & ","
	  end if
	  if rightcol <> "" then
	    rightcol = rightcol & ","
	  end if
	  if mainsticky <> "" then
	    mainsticky = mainsticky & ","
	  end if
	end if
	set fpRS = nothing
	
	sSql = "update " & strTablePrefix & "FP_USERS set fp_leftcol='" & leftcol & "Forum Recent Topics:f_topics_sm',fp_maincol='" & maincol & "Forum News:f_news_fp',fp_rightcol='" & rightcol & "Featured Polls:f_polls_fp' where fp_uid=0"
    my_Conn.execute sSql
	dbHits = dbHits + 1
end sub

sub crArchiveTbls()
  response.Write("<hr /><h4>Create ARCHIVE_TOPICS tables</h4><br />")
':::::::::::::::::::::::::: CREATE ARCHIVE_TOPICS  TABLE :::::::::::::::::::::::::::::::::
sSQL = "CREATE TABLE [" & strTablePrefix & "ARCHIVE_TOPICS]([CAT_ID] LONG, [FORUM_ID] LONG, [ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [T_AUTHOR] LONG, [T_DATE] TEXT(20), [T_IP] TEXT(20), [T_LAST_POST] TEXT(20), [T_LAST_POST_AUTHOR] LONG, [T_LAST_POSTER] LONG, [T_MAIL] LONG, [T_MESSAGE] MEMO, [T_REPLIES] LONG, [T_STATUS] LONG, [T_SUBJECT] TEXT(50), [T_VIEW_COUNT] LONG, [TOPIC_ID] LONG);"

createTable(checkIt(sSQL))

'::::::::::::::::::::::: CREATE ARCHIVE_REPLY  TABLE ::::::::::::::::::::::::::::::::
sSQL = "CREATE TABLE [" & strTablePrefix & "ARCHIVE_REPLY]([CAT_ID] LONG, [FORUM_ID] LONG, [ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [R_AUTHOR] LONG, [R_DATE] TEXT(20), [R_IP] TEXT(20), [R_MAIL] LONG, [R_MESSAGE] MEMO, [REPLY_ID] LONG, [TOPIC_ID] LONG);"

createTable(checkIt(sSQL))
end sub

sub crModeratorTbl()
':::::::::::::::::::::::: CREATE MODERATOR TABLE ::::::::::::::::::::::::::::::::::
  response.Write("<hr /><h4>Create MODERATOR table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "MODERATOR]([FORUM_ID] LONG DEFAULT 0, [MEMBER_ID] LONG DEFAULT 0, [MOD_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [MOD_TYPE] INTEGER DEFAULT 0);"

createTable(checkIt(sSQL))
end sub

sub crAllowedMembers()
'::::::::::::::::::: CREATE " & strTablePrefix & "ALLOWED_MEMBERS TABLE (FORUMS) ::::::::::::::::::::::
  response.Write("<hr /><h4>Create " & strTablePrefix & "ALLOWED_MEMBERS table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "ALLOWED_MEMBERS]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [FORUM_ID] LONG NOT NULL, [MEMBER_ID] LONG NOT NULL);"

createTable(checkIt(sSQL))
end sub
  
sub crCatTbl()
':::::::::::::::::::::::: CREATE CATEGORY TABLE :::::::::::::::::::::::::::::::
  response.Write("<hr /><h4>Create Forum CATEGORY table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "CATEGORY]([CAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CAT_NAME] TEXT(100), [CAT_ORDER] LONG DEFAULT 1, [CAT_STATUS] BYTE DEFAULT 1);"

createTable(checkIt(sSQL))

'-------------------- populate table with default values --------------------------
redim arrData(2)
arrData(0) = "" & strTablePrefix & "CATEGORY"
arrData(1) = "CAT_NAME, CAT_ORDER, CAT_STATUS"
arrData(2) = "'Category', 1, 1"
populateB(arrData)
end sub

sub crForumTbl()
'::::::::::::::::::::: CREATE " & strTablePrefix & "FORUM TABLE :::::::::::::::::::::::::
  response.Write("<hr /><h4>Create category FORUM table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "FORUM]([CAT_ID] LONG DEFAULT 0, [F_COUNT] LONG DEFAULT 0, [F_DESCRIPTION] TEXT(255), [F_IP] TEXT(50) DEFAULT '000.000.000.000', [F_LAST_POST] TEXT(50), [F_LAST_POST_AUTHOR] LONG, [F_MAIL] BYTE DEFAULT 0, [F_PASSWORD_NEW] TEXT(255), [F_PRIVATEFORUMS] LONG DEFAULT 0, [F_STATUS] BYTE DEFAULT 1, [F_SUBJECT] TEXT(100), [F_TOPICS] LONG DEFAULT 0, [F_TYPE] INTEGER DEFAULT 0, [F_URL] TEXT(255), [FORUM_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [FORUM_ORDER] LONG DEFAULT 1, [L_ARCHIVE] TEXT(20));"

createTable(checkIt(sSQL))
		
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", F_STATUS" 
		strSql = strSql & ", F_MAIL" 
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_URL"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TOPICS"
		strSql = strSql & ", F_COUNT"
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_PRIVATEFORUMS"
		strSql = strSql & ", F_TYPE" 
		strSql = strSql & ", F_IP"
		strSql = strSql & ", F_LAST_POST_AUTHOR"
		strSql = strSql & ", F_PASSWORD_NEW"
		strSql = strSql & ", FORUM_ORDER"
		strSql = strSql & ", L_ARCHIVE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'General Chat'"
		strSql = strSql & ", ''"
		strSql = strSql & ", 'General Chat Forum'"
		strSql = strSql & ", 3"
		strSql = strSql & ", 3"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", 1"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 1"
		strSql = strSql & ", ''"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub crTopicTbl()
':::::::::::::::::::::::::: CREATE TOPICS  TABLE :::::::::::::::::::::::::::::::::::::::
  response.Write("<hr /><h4>Create " & strTablePrefix & "TOPICS table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "TOPICS]([CAT_ID] LONG NOT NULL DEFAULT 0, [FORUM_ID] LONG NOT NULL DEFAULT 0, [T_AUTHOR] LONG DEFAULT 0, [T_DATE] TEXT(50), [T_INPLACE] LONG DEFAULT 0, [T_IP] TEXT(50) DEFAULT '000.000.000.000', [T_LAST_POST] TEXT(50), [T_LAST_POST_AUTHOR] LONG, [T_LAST_POSTER] LONG DEFAULT 0, [T_MAIL] BYTE DEFAULT 0, [T_MESSAGE] MEMO, [T_MSGICON] LONG DEFAULT 1, [T_NEWS] LONG DEFAULT 0, [T_POLL] LONG DEFAULT 0, [T_REPLIES] LONG DEFAULT 0, [T_SIG] LONG DEFAULT 1, [T_STATUS] BYTE DEFAULT 1, [T_SUBJECT] TEXT(100), [T_VIEW_COUNT] LONG DEFAULT 0, [TOPIC_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL);"

createTable(checkIt(sSQL))

	'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", T_STATUS"
		strSql = strSql & ", T_MAIL"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_REPLIES" 
		strSql = strSql & ", T_VIEW_COUNT"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_DATE" 
		strSql = strSql & ", T_LAST_POSTER"
		strSql = strSql & ", T_IP"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_MSGICON"
		strSql = strSql & ", T_INPLACE"
		strSql = strSql & ", T_POLL"
		strSql = strSql & ", T_NEWS"
		strSql = strSql & ", T_SIG"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'News Topic'"
		strSql = strSql & ", 'A news topic must be posted in a non-private forum (all visitors).<br /><br /><b>How to post a news topic:</b><br />Log in as Administrator or Moderator<br />Go to the forum you want the News<br />Click on ""New Topic""<br />Write Your ""News""<br />Check the box ""News"", you find it over the button ""Post New Topic""'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", T_STATUS"
		strSql = strSql & ", T_MAIL"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_REPLIES" 
		strSql = strSql & ", T_VIEW_COUNT"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_DATE" 
		strSql = strSql & ", T_LAST_POSTER"
		strSql = strSql & ", T_IP"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_MSGICON"
		strSql = strSql & ", T_INPLACE"
		strSql = strSql & ", T_POLL"
		strSql = strSql & ", T_NEWS"
		strSql = strSql & ", T_SIG"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'Test Topic'"
		strSql = strSql & ", 'Thanks for downloading SkyPortal v" & strVer & ".<br /><br />Please visit our website at www.SkyPortal.net'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", T_STATUS"
		strSql = strSql & ", T_MAIL"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_REPLIES" 
		strSql = strSql & ", T_VIEW_COUNT"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_DATE" 
		strSql = strSql & ", T_LAST_POSTER"
		strSql = strSql & ", T_IP"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_MSGICON"
		strSql = strSql & ", T_INPLACE"
		strSql = strSql & ", T_POLL"
		strSql = strSql & ", T_NEWS"
		strSql = strSql & ", T_SIG"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'Test Poll'"
		strSql = strSql & ", 'Welcome to SkyPortal'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub crReplyTbl()
	'::::::::::::::::::::::::: CREATE REPLY TABLE :::::::::::::::::::::::::::::::::::::::::
  response.Write("<hr /><h4>Create " & strTablePrefix & "REPLY table</h4><br />")
	sSQL = "CREATE TABLE [" & strTablePrefix & "REPLY]([CAT_ID] LONG NOT NULL DEFAULT 0, [FORUM_ID] LONG NOT NULL DEFAULT 0, [R_AUTHOR] LONG DEFAULT 0, [R_DATE] TEXT(50), [R_IP] TEXT(50) DEFAULT '000.000.000.000', [R_MAIL] BYTE DEFAULT 0, [R_MESSAGE] MEMO, [R_MSGICON] LONG DEFAULT 1, [R_SIG] LONG DEFAULT 1, [REPLY_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [TOPIC_ID] LONG NOT NULL DEFAULT 0);"

	createTable(checkIt(sSQL))
end sub

sub crPollTbls()
  response.Write("<hr /><h4>Create POLLS tables</h4><br />")
	'::::::::::::::::::::::: CREATE POLLS TABLE :::::::::::::::::::::::::::::::::::::
	sSQL = "CREATE TABLE [" & strTablePrefix & "POLLS]([ANSWER1] TEXT(50), [ANSWER10] TEXT(50), [ANSWER11] TEXT(50), [ANSWER12] TEXT(50), [ANSWER2] TEXT(50), [ANSWER3] TEXT(50), [ANSWER4] TEXT(50), [ANSWER5] TEXT(50), [ANSWER6] TEXT(50), [ANSWER7] TEXT(50), [ANSWER8] TEXT(50), [ANSWER9] TEXT(50), [END_DATE] TEXT(50), [POLL_ALLOW] LONG DEFAULT 0, [POLL_AUTHOR] LONG DEFAULT 0, [POLL_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [POLL_QUESTION] TEXT(50), [POLL_TYPE] LONG DEFAULT 0, [POST_DATE] TEXT(50), [RESULT1] LONG DEFAULT 0, [RESULT10] LONG DEFAULT 0, [RESULT11] LONG DEFAULT 0, [RESULT12] LONG DEFAULT 0, [RESULT2] LONG DEFAULT 0, [RESULT3] LONG DEFAULT 0, [RESULT4] LONG DEFAULT 0, [RESULT5] LONG DEFAULT 0, [RESULT6] LONG DEFAULT 0, [RESULT7] LONG DEFAULT 0, [RESULT8] LONG DEFAULT 0, [RESULT9] LONG DEFAULT 0);"

	createTable(checkIt(sSQL))

	'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "POLLS "
		strSql = strSql & "(POLL_TYPE"
		strSql = strSql & ", POLL_ALLOW"
		strSql = strSql & ", POLL_QUESTION"
		strSql = strSql & ", ANSWER1"
		strSql = strSql & ", ANSWER2"
		strSql = strSql & ", ANSWER3"
		strSql = strSql & ", ANSWER4"
		strSql = strSql & ", ANSWER5" 
		strSql = strSql & ", ANSWER6"
		strSql = strSql & ", ANSWER7"
		strSql = strSql & ", ANSWER8" 
		strSql = strSql & ", ANSWER9"
		strSql = strSql & ", ANSWER10"
		strSql = strSql & ", ANSWER11"
		strSql = strSql & ", ANSWER12"
		strSql = strSql & ", RESULT1"
		strSql = strSql & ", RESULT2"
		strSql = strSql & ", RESULT3"
		strSql = strSql & ", RESULT4"
		strSql = strSql & ", RESULT5"
		strSql = strSql & ", RESULT6"
		strSql = strSql & ", RESULT7"
		strSql = strSql & ", RESULT8"
		strSql = strSql & ", RESULT9"
		strSql = strSql & ", RESULT10"
		strSql = strSql & ", RESULT11"
		strSql = strSql & ", RESULT12"
		strSql = strSql & ", POST_DATE"
		strSql = strSql & ", END_DATE"
		strSql = strSql & ", POLL_AUTHOR"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'Do you like this site?'"
		strSql = strSql & ", 'Yes'"
		strSql = strSql & ", 'No'"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '20100108000000'"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)

	':::::::::::::::::::::::: CREATE POLL_ANS TABLE ::::::::::::::::::::::::::::::::::::
	sSQL = "CREATE TABLE [" & strTablePrefix & "POLL_ANS]([ANS_DATE] TEXT(50), [ANS_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [ANS_VALUE] TEXT(255), [IP] TEXT(50), [MEMBER_ID] LONG DEFAULT 0, [POLL_ID] LONG DEFAULT 0);"

	createTable(checkIt(sSQL))
end sub
	
sub crReportedTbl()
':::::::::::::::::::: CREATE REPORTED_POST TABLE ::::::::::::::::::::::::::::::::
  response.Write("<hr /><h4>Create REPORTED_POST table</h4><br />")
sSQL = "CREATE TABLE [" & strTablePrefix & "REPORTED_POST]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [R_ACTION_BY] LONG NOT NULL DEFAULT 0, [R_ACTION_DATE] TEXT(50) DEFAULT '0', [R_ACTION_TAKEN] MEMO DEFAULT '0', [R_COMMENTS] MEMO DEFAULT '0', [R_POST] MEMO DEFAULT '0', [R_REASON] MEMO, [R_REPLY_ID] TEXT(21) NOT NULL DEFAULT '0', [R_REPORTED_DATE] TEXT(50) DEFAULT '0', [R_REPORTER_ID] TEXT(21) NOT NULL DEFAULT '0', [R_REPORTER_IP] TEXT(20) DEFAULT '0', [R_STATUS] LONG NOT NULL DEFAULT 0, [R_TOPIC_ID] TEXT(100) NOT NULL DEFAULT '0');"

createTable(checkIt(sSQL))
end sub
	
sub crRelationships()
  response.Write("<hr /><h4>Create RELATIONSHIPS</h4><br />")
'::::::::::::::::::::::: RELATIONSHIPS :::::::::::::::::::::::::::::::::::::::::::
'sSQL = "ALTER TABLE [" & strTablePrefix & "MODERATOR] ADD CONSTRAINT [{2B411B85-40F7-4E14-B275-95D2E831DAD0}] FOREIGN KEY ([FORUM_ID]) REFERENCES [" & strTablePrefix & "FORUM] ([FORUM_ID]) ON UPDATE NO ACTION ON DELETE NO ACTION;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 1<br />")

'sSQL = "ALTER TABLE [" & strTablePrefix & "REPLY] ADD CONSTRAINT [{5F76D07C-3D65-4E78-BBFC-738EBE52886C}] FOREIGN KEY ([R_AUTHOR]) REFERENCES [" & strTablePrefix & "MEMBERS] ([MEMBER_ID]) ON UPDATE NO ACTION ON DELETE NO ACTION;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 2<br />")

'sSQL = "ALTER TABLE [" & strTablePrefix & "TOPICS] ADD CONSTRAINT [{60718A79-B4E2-4A53-B459-668F93D58DFF}] FOREIGN KEY ([T_AUTHOR]) REFERENCES [" & strTablePrefix & "MEMBERS] ([MEMBER_ID]) ON UPDATE NO ACTION ON DELETE NO ACTION;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 3<br />")

'sSQL = "ALTER TABLE [" & strTablePrefix & "FORUM] ADD CONSTRAINT [{7D06D8CC-7202-4B19-A990-C029CFC6CD57}] FOREIGN KEY ([CAT_ID]) REFERENCES [" & strTablePrefix & "CATEGORY] ([CAT_ID]) ON UPDATE CASCADE ON DELETE CASCADE;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 4<br />")

'sSQL = "ALTER TABLE [" & strTablePrefix & "REPLY] ADD CONSTRAINT [{B94F8AC8-B51D-4920-A59B-E41A956CC74E}] FOREIGN KEY ([TOPIC_ID]) REFERENCES [" & strTablePrefix & "TOPICS] ([TOPIC_ID]) ON UPDATE CASCADE ON DELETE CASCADE;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 5<br />")

'sSQL = "ALTER TABLE [" & strTablePrefix & "MODERATOR] ADD CONSTRAINT [{C78D62BA-66DE-44F7-85FD-CF9E4CE16BD9}] FOREIGN KEY ([MEMBER_ID]) REFERENCES [" & strTablePrefix & "MEMBERS] ([MEMBER_ID]) ON UPDATE NO ACTION ON DELETE NO ACTION;"
'my_Conn.execute sSQL
'dbHits = dbHits + 1
'response.Write("Relationship created 6<br />")

    'sSQL = "ALTER TABLE [" & strTablePrefix & "TOPICS] ADD CONSTRAINT [{FFA60B9A-931D-4EE4-A3BF-AA8D4350D507}] FOREIGN KEY ([FORUM_ID]) REFERENCES [" & strTablePrefix & "FORUM] ([FORUM_ID]) ON UPDATE CASCADE ON DELETE CASCADE;"
    'my_Conn.execute sSQL
    'dbHits = dbHits + 1
    'response.Write("Relationship created 7<br />")
end sub

sub uninstall_Forums()
  response.Write("<hr /><h3>Uninstall App</h3><br />")
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
	apid = rsA("APP_ID")
  end if
  set rsA = nothing
  
  sSql = "SELECT * FROM " & strTablePrefix & "FP WHERE APP_ID=" & apid
  set rsA = my_Conn.execute(sSql)
  do while not rsA.eof
	delFPusers(rsA("fp_function"))
	rsA.movenext
  loop
  set rsA = nothing
	
  sSql = "DELETE FROM " & strTablePrefix & "FP WHERE APP_ID=" & apid
  executeThis(sSql)
  
  sSql = "DELETE FROM " & strTablePrefix & "WELCOME WHERE W_MODULE=" & apid
  executeThis(sSql)
 
  sSql = "delete from menu where APP_ID = " & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "PAGES WHERE P_APP=" & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "UPLOAD_CONFIG WHERE UP_APPID=" & apid
  executeThis(sSql)
  
  droptable("" & strTablePrefix & "CATEGORY")
  droptable("" & strTablePrefix & "FORUM")
  droptable("" & strTablePrefix & "TOPICS")
  droptable("" & strTablePrefix & "REPLY")
  droptable("" & strTablePrefix & "REPORTED_POST")
  droptable("" & strTablePrefix & "POLLS")
  droptable("" & strTablePrefix & "POLL_ANS")
  droptable("" & strTablePrefix & "MODERATOR")
  droptable("" & strTablePrefix & "ALLOWED_MEMBERS")
  droptable("" & strTablePrefix & "ARCHIVE_TOPICS")
  droptable("" & strTablePrefix & "ARCHIVE_REPLY")
	
  sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_iNAME='forums'"
  executeThis(sSql)
  mnu.DelMenuFiles("")
  response.Write("<h4>Module Uninstall Complete</h4><br /><hr /><br />")
end sub

sub forum_Upgrade07()
  response.Write("<hr /><h4>Update PORTAL_APPS</h4><br />")
	   sSql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_USERS = '1,2,3'"
	   sSql = sSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
	   sSql = sSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & datetostr(now()) & "'"
	   sSql = sSql & ", APP_SUBSEC = 0 WHERE APP_iNAME = 'forums';"
       executeThis(sSql)
  
  b_forums()
  b_forums_cfg()
end sub

':: FORUM MENU :::::::::::::::::::::::::::::::::::
sub b_forums()
	mnu.DelMenuFiles("")
  response.Write("<hr /><h4>Forums Menu</h4><br />")
  sSql = "select APP_ID from PORTAL_APPS where APP_iNAME = 'forums'"
  set rsT = my_Conn.execute(sSql)
    ap_id = rsT(0)
  set rsT = nothing
  ':: start button template
  mnuName = "* Forums"	
  mnuIName = "b_forums"
  mnuBName = "Forums"

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,iName,mnuFunction,mnuAccess,mnuOrder,app_id"
  arrData(2) = "'" & mnuBName & "', '" & mnuIName & "','','','','','" & mnuName & "','" & mnuIName & "','cntActiveTopics','',1,"& ap_id &""
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and iName = '" & mnuIName & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(5)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onClick,mnuImage,mnuTitle,iName,ParentID,mnuFunction,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'Forum Home', '" & mnuBName & "','fhome.asp','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',1"
  arrData(3) = "'Active Topics', '" & mnuBName & "','forum_active_topics.asp','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'cntActiveTopics',"& ap_id &",'',2"
  arrData(4) = "'Forum Search', '" & mnuBName & "','forum_search.asp','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',3"
  arrData(5) = "'Forum FAQ', '" & mnuBName & "','forum_faq.asp?page=forums','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',4"
  populateB(arrData)
 
 ':: add 'nav_main' menu reference
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,iName,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mnuBName & " Menu', 'nav_main','','','','','Portal Navbar','nav_main',"& ap_id &",'" & mnuIName & "',3"
populateB(arrData)
 
 ':: add 'def_main' menu reference
redim arrData(3)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuFunction,mnuTitle,iName,app_id,mnuAdd,mnuAccess,mnuOrder"
arrData(2) = "'" & mnuBName & " Menu', 'def_main','','','','','','Portal Default','def_main',"& ap_id &",'" & mnuIName & "','',3"
  arrData(3) = "'Reported Posts', 'def_main','forum_report_post_moderate.asp','','','_parent','cntReportedPosts()','Portal Default','def_main',"& ap_id &",'','1',8"
populateB(arrData)
end sub

sub b_forums_cfg()
  response.Write("<hr /><h4>Forums ADMIN Menu</h4><br />")
  sSql = "select APP_ID from PORTAL_APPS where APP_iNAME = 'forums'"
  set rsT = my_Conn.execute(sSql)
    ap_id = rsT(0)
  set rsT = nothing
  ':: start button template
  mnuName = "* Forums ADMIN"	
  mnuIName = "b_forums_cfg"
  mnuBName = "Forums"

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,iName,mnuFunction,mnuAccess,mnuOrder,app_id"
  arrData(2) = "'" & mnuBName & "', '" & mnuIName & "','','','','','" & mnuName & "','" & mnuIName & "','','',1,"& ap_id &""
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and iName = '" & mnuIName & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(11)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onClick,mnuImage,mnuTitle,iName,ParentID,mnuFunction,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'Forum Features', '" & mnuBName & "','admin_forums.asp','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',1"
  arrData(3) = "'Moderator Setup', '" & mnuBName & "','admin_forums.asp?cmd=1','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',2"
  arrData(4) = "'Merge Forums', '" & mnuBName & "','admin_forums.asp?cmd=2','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',3"
  arrData(5) = "'Update Counts', '" & mnuBName & "','admin_count.asp','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',4"
  arrData(6) = "'Forum Archiving', '" & mnuBName & "','admin_forums.asp?cmd=3','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',5"
  arrData(7) = "'Forum Order', '" & mnuBName & "','admin_forums.asp?cmd=4','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',6"
  arrData(8) = "'Forum Status', '" & mnuBName & "','admin_forums.asp?cmd=5','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',7"
  arrData(9) = "'Last Topics', '" & mnuBName & "','admin_forums.asp?cmd=6','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',8"
  arrData(10) = "'Front Page News', '" & mnuBName & "','admin_forums.asp?cmd=7','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',9"
  arrData(11) = "'Polls', '" & mnuBName & "','admin_forums.asp?cmd=8','_parent','','','" & mnuName & "','" & mnuIName & "'," & pID & ",'',"& ap_id &",'',10"
  populateB(arrData)
 
 ':: add module links to module admin menu
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,iName,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mnuBName & " Menu', 'm_admin','','','','','Module Admin','m_admin',"& ap_id &",'" & mnuIName & "',4"
populateB(arrData)

end sub
%>
