<% 
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
%>
<!-- include file="createForums.asp" -->
<%
sub createDB()
	createCore()
	response.Write("<hr /><hr />")
end sub

sub createCore()
	'::this will create the core tables that are for SkyPortal beta
	cr_new_SP_core()
	'::this will create the tables that are new for SkyPortal RC1
	cr_new_SP_RC1()
	'::this will create the tables that are new for SkyPortal RC2
	'::located in create211_SP.asp
	update_rc1_rc2()
	'::this will create the tables that are new for SkyPortal RC3
	'::located in create211_SP.asp
	update_rc2_rc3()
	'::this will update the database for SkyPortal RC4
	'::located in create211_SP.asp
	update_rc3_rc4()
	'::this will update the database for SkyPortal RC5
	'::located in create211_SP.asp
	update_rc4_rc5()
	'::this will update the database for SkyPortal RC6
	'::located in create211_SP.asp
	update_rc5_rc6()
	'::this will update the database for SkyPortal RC7
	'::located in create211_SP.asp
	update_rc6_rc7()
	'::this will update the database for SkyPortal v1.0
	'::located in create211_SP.asp
	update_rc7_v1()
end sub

sub cr_new_SP_core()
	tblAPPS()
	tblConfig()
	tblIPgate()
	tblThemes()
	tblUploads()
	tblBanners()
	tblMemberTables()
	tblAvatars()
	tblCountries()
end sub
	
sub cr_new_SP_RC1()
	'new for SkyPortal RC1
	tblSubscriptions()
	tblBookmarks()
	portalGroups()
	portalGroupMembers()
	'portalGroupPerms()
	portalFrontPage()
	portalFPusers()
	tblAnnounce()
	tblWelcome()
	
end sub

sub tblBookmarks()
'::::::::::::::::::::::::::::: CREATE BOOKMARKS TABLE :::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "BOOKMARKS")
sSQL = "CREATE TABLE [" & strTablePrefix & "BOOKMARKS]([BOOKMARK_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [APP_ID] LONG NOT NULL, [M_ID] LONG NOT NULL, [CAT_ID] LONG NULL, [SUBCAT_ID] LONG NULL, [ITEM_ID] LONG NULL, [ITEM_URL] MEMO NULL, [ITEM_TITLE] MEMO NULL);"

createTable(checkIt(sSQL))

redim indexes(1)
indexes(0) = "CREATE INDEX [M_ID] ON [" & strTablePrefix & "BOOKMARKS]([M_ID]);"
indexes(1) = "CREATE INDEX [ITEM_ID] ON [" & strTablePrefix & "BOOKMARKS]([ITEM_ID]);"
indexes(1) = "CREATE INDEX [APP_ID] ON [" & strTablePrefix & "BOOKMARKS]([APP_ID]);"
createIndx(indexes)
end sub

sub tblSubscriptions()
'::::::::::::::::::::: CREATE " & strTablePrefix & "SUBSCRIPTIONS TABLE ::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "SUBSCRIPTIONS")
sSQL = "CREATE TABLE [" & strTablePrefix & "SUBSCRIPTIONS]([SUBSCRIPTION_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [APP_ID] LONG NOT NULL, [M_ID] LONG NOT NULL, [CAT_ID] LONG NULL, [SUBCAT_ID] LONG NULL, [ITEM_ID] LONG NULL, [ITEM_URL] MEMO NULL, [ITEM_TITLE] MEMO NULL);"

createTable(checkIt(sSQL))

redim indexes(2)
indexes(0) = "CREATE INDEX [APP_ID] ON [" & strTablePrefix & "SUBSCRIPTIONS]([APP_ID]);"
indexes(1) = "CREATE INDEX [M_ID] ON [" & strTablePrefix & "SUBSCRIPTIONS]([M_ID]);"
indexes(2) = "CREATE INDEX [ITEM_ID] ON [" & strTablePrefix & "SUBSCRIPTIONS]([ITEM_ID]);"
createIndx(indexes)
end sub

sub tblAnnounce()
'::::::::::::::::::: CREATE " & strTablePrefix & "ANNOUNCEMENTS TABLE :::::::::::::::::::::::::
droptable("" & strTablePrefix & "ANNOUNCEMENTS")
sSQL = "CREATE TABLE [" & strTablePrefix & "ANNOUNCEMENTS]([A_ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, [A_AUTHOR] INT, [A_SUBJECT] TEXT(200), [A_MESSAGE] MEMO, [A_START_DATE] TEXT(50), [A_END_DATE] TEXT(50));"

createTable(checkIt(sSQL))
'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "ANNOUNCEMENTS "
		strSql = strSql & "(A_AUTHOR"
		strSql = strSql & ", A_SUBJECT"
		strSql = strSql & ", A_MESSAGE"
		strSql = strSql & ", A_START_DATE"
		strSql = strSql & ", A_END_DATE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1"
		strSql = strSql & ", '" & txtSUPostInst & "'"
		strSql = strSql & ", '" & txtSUPIDesc & "'"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", '" & DateToStr(dateAdd("d", 30, now())) & "'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub tblWelcome()
'::::::::::::::::::: CREATE " & strTablePrefix & "Welcome  TABLE :::::::::::::::::::::::::
droptable("" & strTablePrefix & "WELCOME")
sSQL = "CREATE TABLE [" & strTablePrefix & "WELCOME]([W_ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, [W_TITLE] TEXT(200), [W_SUBJECT] TEXT(200), [W_MESSAGE] MEMO, [W_DELETE] INT, [W_ACTIVE] INT, [W_MODULE] INT);"

createTable(checkIt(sSQL))
'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "WELCOME "
		strSql = strSql & "(W_TITLE"
		strSql = strSql & ", W_SUBJECT"
		strSql = strSql & ", W_MESSAGE"
		strSql = strSql & ", W_DELETE"
		strSql = strSql & ", W_MODULE"
		strSql = strSql & ", W_ACTIVE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txtSUWelMsg & "'"
		strSql = strSql & ", '" & txtSUWelMsgTitle & "'"
		strSql = strSql & ", '" & txtSUWelMsgDesc & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "WELCOME "
		strSql = strSql & "(W_TITLE"
		strSql = strSql & ", W_SUBJECT"
		strSql = strSql & ", W_MESSAGE"
		strSql = strSql & ", W_DELETE"
		strSql = strSql & ", W_MODULE"
		strSql = strSql & ", W_ACTIVE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txtSUWelMsg2 & "'"
		strSql = strSql & ", '" & txtSUWelMsgTitle2 & "'"
		strSql = strSql & ", '" & txtSUWelMsgDesc & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub tblAvatars()
':::::::::::::::::::::::: CREATE AVATAR TABLE ::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "AVATAR")
sSQL = "CREATE TABLE [" & strTablePrefix & "AVATAR]([A_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [A_MEMBER_ID] LONG DEFAULT 0, [A_NAME] TEXT(50), [A_URL] TEXT(255));"

createTable(checkIt(sSQL))

redim indexes(0)
indexes(0) = "CREATE INDEX [A_MEMBER_ID] ON [" & strTablePrefix & "AVATAR]([A_MEMBER_ID]);"
createIndx(indexes)

'-------------------- populate table with default values --------------------------
redim arrData(17)
arrData(0) = "" & strTablePrefix & "AVATAR"
arrData(1) = "A_NAME, A_URL"
arrData(2) = "'noavatar', 'files/avatars/noavatar.gif'"
arrData(3) = "'120', 'files/avatars/120.jpg'"
arrData(4) = "'124', 'files/avatars/124.jpg'"
arrData(5) = "'214', 'files/avatars/214.jpg'"
arrData(6) = "'226', 'files/avatars/226a.jpg'"
arrData(7) = "'256', 'files/avatars/256.jpg'"
arrData(8) = "'264', 'files/avatars/264.jpg'"
arrData(9) = "'269', 'files/avatars/269.jpg'"
arrData(10) = "'271', 'files/avatars/271.jpg'"
arrData(11) = "'315', 'files/avatars/315.jpg'"
arrData(12) = "'318', 'files/avatars/318.jpg'"
arrData(13) = "'320', 'files/avatars/320.jpg'"
arrData(14) = "'321', 'files/avatars/321.jpg'"
arrData(15) = "'322', 'files/avatars/322.jpg'"
arrData(16) = "'323', 'files/avatars/323.jpg'"
arrData(17) = "'489', 'files/avatars/489.jpg'"
populateB(arrData)

':::::::::::::::::::::::::::::::::: CREATE AVATAR2 TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "AVATAR2")
sSQL = "CREATE TABLE [" & strTablePrefix & "AVATAR2]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [A_BORDER] LONG, [A_HSIZE] LONG, [A_WSIZE] LONG);"

createTable(checkIt(sSQL))

'-------------------- populate table with default values --------------------------
redim arrData(2)
arrData(0) = "" & strTablePrefix & "AVATAR2 "
arrData(1) = "A_HSIZE, A_WSIZE, A_BORDER"
arrData(2) = "64, 64, 0"
populateB(arrData)
end sub 'tblAvatars

sub tblBanners()
':::::::::::::::::::::::: CREATE BANNERS TABLE ::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "BANNERS")
sSQL = "CREATE TABLE [" & strTablePrefix & "BANNERS]([B_ACRONYM] TEXT(100) NOT NULL, [B_ACTIVATED_DATE] TEXT(255) NOT NULL, [B_ACTIVE] BYTE NOT NULL DEFAULT 1, [B_HITS] LONG NOT NULL DEFAULT 0, [B_IMAGE] TEXT(255) NOT NULL, [B_IMPRESSIONS] LONG NOT NULL DEFAULT 0, [B_LINKTO] TEXT(100) NOT NULL, [B_LOCATION] BYTE NOT NULL DEFAULT 1, [B_NAME] TEXT(50) NOT NULL, [ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL);"

createTable(checkIt(sSQL))

'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'SkyPortal v" &  strVer & "'"
		strSql = strSql & ", 'http://www.SkyPortal.net'"
		strSql = strSql & ", '" & txtSUBan1 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/SkyPortal.gif'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'WebDogg Hosting'"
		strSql = strSql & ", 'http://www.webdogg.com'"
		strSql = strSql & ", '" & txtSUBan2 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/webdogg.jpg'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'LiveAir Networks'"
		strSql = strSql & ", 'https://www.securepaynet.net/gdshop/rhp/default.asp?prog_id=lnwgoodies'"
		strSql = strSql & ", '" & txtSUBan3 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/liveair.gif'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'SkyPortal'"
		strSql = strSql & ", 'http://www.SkyPortal.net'"
		strSql = strSql & ", '" & txtSUBan4 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/aff_SkyPortal.gif'"
		strSql = strSql & ", 2"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'WebDogg Hosting'"
		strSql = strSql & ", 'http://www.webdogg.com'"
		strSql = strSql & ", '" & txtSUBan2 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/aff_webdogg.gif'"
		strSql = strSql & ", 2"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "BANNERS "
		strSql = strSql & "(B_NAME"
		strSql = strSql & ", B_LINKTO"
		strSql = strSql & ", B_ACRONYM"
		strSql = strSql & ", B_HITS"
		strSql = strSql & ", B_ACTIVE"
		strSql = strSql & ", B_ACTIVATED_DATE"
		strSql = strSql & ", B_IMAGE"
		strSql = strSql & ", B_LOCATION"
		strSql = strSql & ", B_IMPRESSIONS" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'LiveAir Networks'"
		strSql = strSql & ", 'https://www.securepaynet.net/gdshop/rhp/default.asp?prog_id=lnwgoodies'"
		strSql = strSql & ", '" & txtSUBan5 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", 'files/banners/aff_liveair.gif'"
		strSql = strSql & ", 2"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub tblThemes()
':::::::::::::::::::: CREATE COLORS TABLE :::::::::::::::::::::::::::
droptable("" & strTablePrefix & "COLORS")
sSQL = "CREATE TABLE [" & strTablePrefix & "COLORS]([C_STRAUTHOR] TEXT(200), [C_STRDESCRIPTION] TEXT(255), [C_STRFOLDER] TEXT(50), [C_TEMPLATE] TEXT(50), [C_STRTITLEIMAGE] TEXT(50), [C_INTSUBSKIN] INTEGER, [CONFIG_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [C_SKINLEVEL] TEXT(255));"

createTable(checkIt(sSQL))
		
		strSql = "INSERT INTO " & strTablePrefix & "COLORS "
		strSql = strSql & "(C_STRFOLDER"
		strSql = strSql & ", C_STRDESCRIPTION"
		strSql = strSql & ", C_STRAUTHOR"
		strSql = strSql & ", C_TEMPLATE"
		strSql = strSql & ", C_STRTITLEIMAGE"
		strSql = strSql & ", C_INTSUBSKIN"
		strSql = strSql & ", C_SKINLEVEL"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & itFolder & "'"
		strSql = strSql & ", '" & itDesc & "'"
		strSql = strSql & ", '" & itAuthor & "'"
		strSql = strSql & ", '" & itName & "'"
		strSql = strSql & ", '" & itLogo & "'"
		strSql = strSql & ", " & itSubSkin
		strSql = strSql & ", '1,2,3'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub tblConfig()
'::::::::::::::::::::: CREATE " & strTablePrefix & "CONFIG  TABLE :::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "CONFIG")
sSQL = "CREATE TABLE [" & strTablePrefix & "CONFIG]([AUTOPM_MESSAGE] MEMO, [AUTOPM_ON] LONG DEFAULT 1, [AUTOPM_SUBJECTLINE] TEXT(255), [C_ALLOWUPLOADS] BYTE DEFAULT 0, [C_DOWNMSG] TEXT(255), [C_FEATUREDPOLL] LONG DEFAULT 0, [C_FORUMSTATUS] TEXT(50), [C_FORUMSUBSCRIPTION] LONG DEFAULT 1, [C_ICALEXIST] LONG DEFAULT 0, [C_ICALNEW] LONG DEFAULT 0, [C_INTHOTTOPICNUM] LONG DEFAULT 0, [C_INTRANKLEVEL0] INTEGER DEFAULT 0, [C_INTRANKLEVEL1] INTEGER DEFAULT 0, [C_INTRANKLEVEL2] INTEGER DEFAULT 0, [C_INTRANKLEVEL3] INTEGER DEFAULT 0, [C_INTRANKLEVEL4] INTEGER DEFAULT 0, [C_INTRANKLEVEL5] INTEGER DEFAULT 0, [C_JOKEOFTHEWEEK] LONG DEFAULT 0, [C_PMTYPE] BYTE DEFAULT 2, [C_POLLCREATE] LONG DEFAULT 0, [C_REMINDERS] LONG DEFAULT 0, [C_STRAGE] INTEGER, [C_STRAIM] BYTE DEFAULT 1, [C_STRALLOWFORUMCODE] BYTE DEFAULT 1, [C_STRALLOWHTML] BYTE DEFAULT 0, [C_STRAUTHTYPE] TEXT(50), [C_STRAUTOLOGON] INTEGER, [C_STRBADWORDFILTER] BYTE DEFAULT 1, [C_STRBADWORDS] TEXT(255), [C_STRBIO] INTEGER, [C_STRCITY] INTEGER, [C_STRCOPYRIGHT] TEXT(200), [C_STRCOUNTRY] INTEGER, [C_STRDATETYPE] TEXT(50), [C_STRDEFTHEME] TEXT(50), [C_STREDITEDBYDATE] BYTE DEFAULT 1, [C_STREMAIL] BYTE DEFAULT 0, [C_STREMAILVAL] INTEGER DEFAULT 0, [C_STRFAVLINKS] INTEGER, [C_STRFLOODCHECK] LONG DEFAULT 0, [C_STRFLOODCHECKTIME] LONG DEFAULT -30, [C_STRFULLNAME] INTEGER, [C_STRGFXBUTTONS] BYTE DEFAULT 1, [C_STRGLOW] BYTE DEFAULT 1, [C_STRHIDEEMAIL] BYTE DEFAULT 0, [C_STRHOBBIES] INTEGER, [C_STRHOMEPAGE] BYTE DEFAULT 1, [C_STRHOMEURL] TEXT(255), [C_STRHOTTOPIC] BYTE DEFAULT 1, [C_STRICONS] BYTE DEFAULT 1, [C_STRICQ] BYTE DEFAULT 1, [C_STRICSLOCATION] TEXT(50), [C_STRIMGINPOSTS] BYTE DEFAULT 0, [C_STRIPGATEBAN] TEXT(2), [C_STRIPGATECOK] TEXT(2), [C_STRIPGATECSS] TEXT(2), [C_STRIPGATEEXP] TEXT(3), [C_STRIPGATELCK] TEXT(2), [C_STRIPGATELKMSG] TEXT(100), [C_STRIPGATELOG] TEXT(2), [C_STRIPGATEMET] TEXT(2), [C_STRIPGATEMSG] TEXT(100), [C_STRIPGATENOACMSG] TEXT(100), [C_STRIPGATETYP] TEXT(2), [C_STRIPGATEVER] TEXT(15), [C_STRIPGATEWARNMSG] TEXT(100), [C_STRIPLOGGING] BYTE DEFAULT 1, [C_STRLNEWS] INTEGER, [C_STRLOGINTYPE] BYTE DEFAULT 0, [C_STRLOGONFORMAIL] INTEGER, [C_STRMAILMODE] TEXT(100), [C_STRMAILSERVER] TEXT(255), [C_STRMARSTATUS] INTEGER, [C_STRMOVETOPICMODE] BYTE DEFAULT 1, [C_STRMSN] BYTE DEFAULT 1, [C_STRNEWREG] LONG DEFAULT 1, [C_STRNTGROUPS] INTEGER, [C_STROCCUPATION] INTEGER, [C_STRPAGENUMBERSIZE] INTEGER, [C_STRPAGESIZE] INTEGER, [C_STRPICTURE] INTEGER, [C_STRPRIVATEFORUMS] BYTE DEFAULT 0, [C_STRQUICKREPLY] LONG DEFAULT 0, [C_STRQUOTE] INTEGER, [C_STRRANKADMIN] TEXT(50), [C_STRRANKCOLOR0] TEXT(50), [C_STRRANKCOLOR1] TEXT(50), [C_STRRANKCOLOR2] TEXT(50), [C_STRRANKCOLOR3] TEXT(50), [C_STRRANKCOLOR4] TEXT(50), [C_STRRANKCOLOR5] TEXT(50), [C_STRRANKCOLORADMIN] TEXT(50), [C_STRRANKCOLORMOD] TEXT(50), [C_STRRANKLEVEL0] TEXT(50), [C_STRRANKLEVEL1] TEXT(50), [C_STRRANKLEVEL2] TEXT(50), [C_STRRANKLEVEL3] TEXT(50), [C_STRRANKLEVEL4] TEXT(50), [C_STRRANKLEVEL5] TEXT(50), [C_STRRANKMOD] TEXT(50), [C_STRRECENTTOPICS] INTEGER, [C_STRSECUREADMIN] BYTE DEFAULT 1, [C_STRSENDER] TEXT(255), [C_STRSEX] INTEGER, [C_STRSHOWIMAGEPOWEREDBY] INTEGER, [C_STRSHOWMODERATORS] BYTE DEFAULT 1, [C_STRSHOWPAGING] INTEGER, [C_STRSHOWRANK] BYTE DEFAULT 0, [C_STRSHOWSTATISTICS] INTEGER, [C_STRSHOWTOPICNAV] INTEGER, [C_STRSIGNATURES] BYTE DEFAULT 1, [C_STRSITETITLE] TEXT(255), [C_STRSTATE] INTEGER, [C_STRTIMEADJUST] LONG DEFAULT 0, [C_STRTIMEADJUSTLOCATION] TEXT(50), [C_STRTIMETYPE] TEXT(50), [C_STRTITLEIMAGE] TEXT(255), [C_STRUNIQUEEMAIL] BYTE DEFAULT 1, [C_STRVAR1] TEXT(50), [C_STRVAR2] TEXT(50), [C_STRVAR3] TEXT(50), [C_STRVAR4] TEXT(50), [C_STRLOCKDOWN] BYTE, [C_STRYAHOO] BYTE DEFAULT 1, [C_STRZIP] BYTE DEFAULT 1, [C_SECIMAGE] INTEGER, [C_STRHEADERTYPE] LONG DEFAULT 0, [CONFIG_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [C_STRVAR8] TEXT(50), [C_STRVAR9] TEXT(50), [C_INTSUBSKIN] INTEGER, [C_VERSION] TEXT(20), [C_ONEADAYDATE] TEXT(20));"
createTable(checkIt(sSQL)) 

	'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "CONFIG "
		strSql = strSql & "(C_STRSITETITLE"
		strSql = strSql & ", C_STRCOPYRIGHT"
		strSql = strSql & ", C_STRTITLEIMAGE"
		strSql = strSql & ", C_STRLOCKDOWN"
		strSql = strSql & ", C_STRHOMEURL"
		strSql = strSql & ", C_STRAUTHTYPE"
		strSql = strSql & ", C_STREMAIL" 
		strSql = strSql & ", C_STRUNIQUEEMAIL"
		strSql = strSql & ", C_STRMAILMODE"
		strSql = strSql & ", C_STRMAILSERVER" 
		strSql = strSql & ", C_STRSENDER"
		strSql = strSql & ", C_STRDATETYPE"
		strSql = strSql & ", C_STRTIMETYPE"
		strSql = strSql & ", C_STRTIMEADJUSTLOCATION"
		strSql = strSql & ", C_STRTIMEADJUST"
		strSql = strSql & ", C_STRMOVETOPICMODE"
		strSql = strSql & ", C_STRPRIVATEFORUMS"
		strSql = strSql & ", C_STRSHOWMODERATORS"
		strSql = strSql & ", C_STRSHOWRANK"
		strSql = strSql & ", C_STRHIDEEMAIL"
		strSql = strSql & ", C_STRIPLOGGING"
		strSql = strSql & ", C_STRALLOWFORUMCODE"
		strSql = strSql & ", C_STRIMGINPOSTS"
		strSql = strSql & ", C_STRALLOWHTML"
		strSql = strSql & ", C_STREDITEDBYDATE"
		strSql = strSql & ", C_STRHOTTOPIC"
		strSql = strSql & ", C_INTHOTTOPICNUM"
		strSql = strSql & ", C_STRHOMEPAGE"
		strSql = strSql & ", C_STRAIM"
		strSql = strSql & ", C_STRYAHOO"
		strSql = strSql & ", C_STRMSN"
		strSql = strSql & ", C_STRICQ"
		strSql = strSql & ", C_STRICONS"
		strSql = strSql & ", C_STRBADWORDFILTER"
		strSql = strSql & ", C_STRBADWORDS"
		strSql = strSql & ", C_STRRANKADMIN"
		strSql = strSql & ", C_STRRANKMOD"
		strSql = strSql & ", C_STRRANKLEVEL0"
		strSql = strSql & ", C_STRRANKLEVEL1"
		strSql = strSql & ", C_STRRANKLEVEL2"
		strSql = strSql & ", C_STRRANKLEVEL3"
		strSql = strSql & ", C_STRRANKLEVEL4"
		strSql = strSql & ", C_STRRANKLEVEL5"
		strSql = strSql & ", C_STRRANKCOLORADMIN"
		strSql = strSql & ", C_STRRANKCOLORMOD"
		strSql = strSql & ", C_STRRANKCOLOR0"
		strSql = strSql & ", C_STRRANKCOLOR1"
		strSql = strSql & ", C_STRRANKCOLOR2"
		strSql = strSql & ", C_STRRANKCOLOR3"
		strSql = strSql & ", C_STRRANKCOLOR4"
		strSql = strSql & ", C_STRRANKCOLOR5"
		strSql = strSql & ", C_INTRANKLEVEL0"
		strSql = strSql & ", C_INTRANKLEVEL1"
		strSql = strSql & ", C_INTRANKLEVEL2"
		strSql = strSql & ", C_INTRANKLEVEL3"
		strSql = strSql & ", C_INTRANKLEVEL4"
		strSql = strSql & ", C_INTRANKLEVEL5"
		strSql = strSql & ", C_STRSIGNATURES"
		strSql = strSql & ", C_STRSHOWSTATISTICS"
		strSql = strSql & ", C_STRSHOWIMAGEPOWEREDBY"
		strSql = strSql & ", C_STRLOGONFORMAIL"
		strSql = strSql & ", C_STRSHOWPAGING"
		strSql = strSql & ", C_STRSHOWTOPICNAV"
		strSql = strSql & ", C_STRPAGESIZE"
		strSql = strSql & ", C_STRPAGENUMBERSIZE"
		strSql = strSql & ", C_STRFULLNAME"
		strSql = strSql & ", C_STRPICTURE"
		strSql = strSql & ", C_STRSEX"
		strSql = strSql & ", C_STRCITY"
		strSql = strSql & ", C_STRSTATE"
		strSql = strSql & ", C_STRAGE"
		strSql = strSql & ", C_STRCOUNTRY"
		strSql = strSql & ", C_STROCCUPATION"
		strSql = strSql & ", C_STRBIO"
		strSql = strSql & ", C_STRHOBBIES"
		strSql = strSql & ", C_STRLNEWS"
		strSql = strSql & ", C_STRQUOTE"
		strSql = strSql & ", C_STRMARSTATUS"
		strSql = strSql & ", C_STRFAVLINKS"
		strSql = strSql & ", C_STRRECENTTOPICS"
		strSql = strSql & ", C_STRAUTOLOGON"
		strSql = strSql & ", C_STREMAILVAL"
		strSql = strSql & ", C_STRNTGROUPS"
		strSql = strSql & ", C_FORUMSTATUS"
		strSql = strSql & ", C_DOWNMSG"
		strSql = strSql & ", C_STRFLOODCHECK"
		strSql = strSql & ", C_STRFLOODCHECKTIME"
		strSql = strSql & ", C_JOKEOFTHEWEEK"
		strSql = strSql & ", C_STRNEWREG"
		strSql = strSql & ", C_POLLCREATE"
		strSql = strSql & ", C_FEATUREDPOLL"
		strSql = strSql & ", C_STRQUICKREPLY"
		strSql = strSql & ", C_STRDEFTHEME"
		strSql = strSql & ", C_PMTYPE"
		strSql = strSql & ", C_ALLOWUPLOADS"
		strSql = strSql & ", C_STRICSLOCATION"
		strSql = strSql & ", C_REMINDERS"
		strSql = strSql & ", C_ICALEXIST"
		strSql = strSql & ", C_ICALNEW"
		strSql = strSql & ", C_STRVAR1"
		strSql = strSql & ", C_STRVAR2"
		strSql = strSql & ", C_STRVAR3"
		strSql = strSql & ", C_STRVAR4"
		strSql = strSql & ", C_FORUMSUBSCRIPTION"
		strSql = strSql & ", AUTOPM_ON"
		strSql = strSql & ", AUTOPM_SUBJECTLINE"
		strSql = strSql & ", AUTOPM_MESSAGE"
		strSql = strSql & ", C_STRZIP"
		strSql = strSql & ", C_STRIPGATEBAN"
		strSql = strSql & ", C_STRIPGATELCK"
		strSql = strSql & ", C_STRIPGATECOK"
		strSql = strSql & ", C_STRIPGATEMET"
		strSql = strSql & ", C_STRIPGATEMSG"
		strSql = strSql & ", C_STRIPGATELKMSG"
		strSql = strSql & ", C_STRIPGATENOACMSG"
		strSql = strSql & ", C_STRIPGATEWARNMSG"
		strSql = strSql & ", C_STRIPGATEVER"
		strSql = strSql & ", C_STRIPGATELOG"
		strSql = strSql & ", C_STRIPGATETYP"
		strSql = strSql & ", C_STRIPGATEEXP"
		strSql = strSql & ", C_STRIPGATECSS"
		strSql = strSql & ", C_STRLOGINTYPE"
		strSql = strSql & ", C_STRHEADERTYPE"
		strSql = strSql & ", C_STRGLOW"
		strSql = strSql & ", C_SECIMAGE"
		strSql = strSql & ", C_STRVAR8"
		strSql = strSql & ", C_STRVAR9"
		strSql = strSql & ", C_INTSUBSKIN"
		strSql = strSql & ", C_ONEADAYDATE"
		strSql = strSql & ", C_VERSION"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & siteName & "'"
		strSql = strSql & ", '" & replace(txtSUAllRtsRes,"[%sitename%]",sitename) & "'"
		strSql = strSql & ", 'site_Logo.jpg'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & portalUrl & "'"
		strSql = strSql & ", 'db'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & emailComponent & "'"
		strSql = strSql & ", '" & mailServer & "'"
		strSql = strSql & ", '" & emailAddy & "'"
		strSql = strSql & ", 'mdy'"
		strSql = strSql & ", '12'"
		strSql = strSql & ", 'server'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 3"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 15"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & txtSUBadWrds & "'"
		strSql = strSql & ", '" & txtSUAdmin & "'"
		strSql = strSql & ", '" & txtSUModerator & "'"
		strSql = strSql & ", '" & txtSUTitle1 & "'"
		strSql = strSql & ", '" & txtSUTitle2 & "'"
		strSql = strSql & ", '" & txtSUTitle3 & "'"
		strSql = strSql & ", '" & txtSUTitle4 & "'"
		strSql = strSql & ", '" & txtSUTitle5 & "'"
		strSql = strSql & ", '" & txtSUTitle6 & "'"
		strSql = strSql & ", 'gold'"
		strSql = strSql & ", 'silver'"
		strSql = strSql & ", 'bronze'"
		strSql = strSql & ", 'orange'"
		strSql = strSql & ", 'cyan'"
		strSql = strSql & ", 'blue'"
		strSql = strSql & ", 'purple'"
		strSql = strSql & ", 'red'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 50"
		strSql = strSql & ", 150"
		strSql = strSql & ", 500"
		strSql = strSql & ", 1200"
		strSql = strSql & ", 2500"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 15"
		strSql = strSql & ", 10"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 5"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'up'"
		strSql = strSql & ", '" & txtSUForumDwn & "'"
		strSql = strSql & ", 1"
		strSql = strSql & ", -30"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 2"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & installTheme & "'"
		strSql = strSql & ", 2"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'files/eventfile.ics'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & txtSUVar1 & "'"
		strSql = strSql & ", '" & txtSUVar2 & "'"
		strSql = strSql & ", '" & txtSUVar3 & "'"
		strSql = strSql & ", '" & txtSUVar4 & "'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & txtSUPMWelcome & "'"
		strSql = strSql & ", '" & txtSUPMWelcomeMsg & "'"
		strSql = strSql & ", 1"
		strSql = strSql & ", '0'"
		strSql = strSql & ", '0'"
		strSql = strSql & ", '1'"
		strSql = strSql & ", '1'"
		strSql = strSql & ", '" & txtSUBanned & "'"
		strSql = strSql & ", '" & txtSUFrmLckd & "'"
		strSql = strSql & ", '" & txtSUNoAccess & "'"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 'Ver 2.4.0'"
		strSql = strSql & ", '1'"
		strSql = strSql & ", '0'"
		strSql = strSql & ", '15'"
		strSql = strSql & ", '0'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 2"
		strSql = strSql & ", 1"	
		strSql = strSql & ", 0"
		strSql = strSql & ", '1'"
		strSql = strSql & ", '0'"
		strSql = strSql & ", 1"
		strSql = strSql & ", '" & date() & "'"	'C_ONEADAYDATE
		strSql = strSql & ", '" & longVer & "'"					
		strSql = strSql & ")"
		populateA(strSql)
		'response.Write("<br />" & strSql & "<br />")

	':::::::::::::::::::::::: CREATE ONLINE TABLE ::::::::::::::::::::::::::::::::::
	droptable("" & strTablePrefix & "ONLINE")
	sSQL = "CREATE TABLE [" & strTablePrefix & "ONLINE]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CheckedIn] TEXT(100), [DateCreated] TEXT(100), [LastChecked] TEXT(100), [LastDateChecked] TEXT(100), [M_BROWSE] MEMO, [UserID] TEXT(100), [UserIP] TEXT(255), [UserAgent] TEXT(100));"

	createTable(checkIt(sSQL))

	createIndex("CREATE INDEX [UserID] ON [" & strTablePrefix & "ONLINE]([UserID]);")

'::::::::::::::::::::: CREATE MODS TABLE ::::::::::::::::::::::::::::::::::::
	droptable("" & strTablePrefix & "MODS")
	sSQL = "CREATE TABLE [" & strTablePrefix & "MODS]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [M_CODE] TEXT(20) NOT NULL, [M_NAME] TEXT(20) NOT NULL, [M_VALUE] TEXT(30) NOT NULL);"

	createTable(checkIt(sSQL))

	createIndex("CREATE INDEX [M_CODE] ON [" & strTablePrefix & "MODS]([M_CODE]);")

	redim arrData(13)
	arrData(0) = "" & strTablePrefix & "MODS"
	arrData(1) = "M_NAME, M_CODE, M_VALUE"
	arrData(2) = "'news', 'slColumns', '1'"
	arrData(3) = "'news', 'slDefimg', 'images/news.gif'"
	arrData(9) = "'slash', 'slEncode', '0'"
	arrData(4) = "'news', 'slEncode', '1'"
	arrData(10) = "'slash', 'slImages', '1'"
	arrData(5) = "'news', 'slImages', '0'"
	arrData(11) = "'slash', 'slLength', '0'"
	arrData(6) = "'news', 'slLength', '270'"
	arrData(12) = "'slash', 'slPosts', '5'"
	arrData(7) = "'news', 'slPosts', '6'"
	arrData(13) = "'slash', 'slSort', '2'"
	arrData(8) = "'news', 'slSort', '1'"
	populateB(arrData)

'::::::::::::::::::::::: CREATE SPAM TABLE :::::::::::::::::::::::::::::::::::::::::
	droptable("" & strTablePrefix & "SPAM")
	sSQL = "CREATE TABLE [" & strTablePrefix & "SPAM]([ARCHIVE] TEXT(1) DEFAULT 0, [F_SENT] TEXT(255), [ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [MESSAGE] MEMO, [SUBJECT] TEXT(255));"

	createTable(checkIt(sSQL))

':::::::::::::::::::::::: CREATE TOTALS TABLE :::::::::::::::::::::::::::::::::::
	droptable("" & strTablePrefix & "TOTALS")
	sSQL = "CREATE TABLE [" & strTablePrefix & "TOTALS]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [COUNT_ID] INTEGER NOT NULL DEFAULT 0, [P_COUNT] LONG NOT NULL DEFAULT 0, [T_COUNT] LONG NOT NULL DEFAULT 0, [U_COUNT] LONG NOT NULL DEFAULT 0);"

	createTable(checkIt(sSQL))

	createIndex("CREATE INDEX [COUNT_ID] ON [" & strTablePrefix & "TOTALS]([COUNT_ID]);")

	redim arrData(2)
	arrData(0) = "" & strTablePrefix & "TOTALS"
	arrData(1) = "P_COUNT, T_COUNT, U_COUNT"
	arrData(2) = "3, 3, 1"
	populateB(arrData)
end sub 'tblConfig

sub tblIPgate()
'::::::::::::::::::::::: CREATE IPLIST  TABLE ::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "IPLIST")
sSQL = "CREATE TABLE [" & strTablePrefix & "IPLIST]([IPLIST_COMMENT] TEXT(255), [IPLIST_DBPAGEKEY] TEXT(32), [IPLIST_ENDDATE] TEXT(32), [IPLIST_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [IPLIST_MEMBERID] TEXT(32) DEFAULT 0, [IPLIST_STARTDATE] TEXT(32), [IPLIST_STARTIP] TEXT(32), [IPLIST_STATUS] TEXT(8));"

createTable(checkIt(sSQL))

'::::::::::::::::::::::: CREATE IPLOG  TABLE ::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "IPLOG")
sSQL = "CREATE TABLE [" & strTablePrefix & "IPLOG]([IPLOG_DATE] TEXT(32), [IPLOG_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [IPLOG_IP] TEXT(32), [IPLOG_MEMBERID] TEXT(32) DEFAULT 0, [IPLOG_PATHINFO] TEXT(255));"

createTable(checkIt(sSQL))

'::::::::::::::::::::: CREATE " & strTablePrefix & "PAGEKEYS TABLE - ipGate :::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "PAGEKEYS")
sSQL = "CREATE TABLE [" & strTablePrefix & "PAGEKEYS]([PAGEKEYS_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [PAGEKEYS_PAGEKEY] TEXT(32));"

createTable(checkIt(sSQL))

redim arrData(4)
arrData(0) = "" & strTablePrefix & "PAGEKEYS"
arrData(1) = "PAGEKEYS_PAGEKEY"
arrData(2) = "'fhome.asp'"
arrData(3) = "'admin_login.asp'"
arrData(4) = "'default.asp'"
populateB(arrData)
end sub

sub tblMemberTables()
':::::::::::::::::::: CREATE MEMBERS  TABLE :::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "MEMBERS")
sSQL = "CREATE TABLE [" & strTablePrefix & "MEMBERS]([MEMBER_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL,[M_AGE] TEXT(10), [M_AIM] TEXT(150), [M_AVATAR_URL] TEXT(255), [M_BIO] MEMO, [M_CITY] TEXT(100), [M_COUNTRY] TEXT(40), [M_DATE] TEXT(50), [M_DEFAULT_VIEW] LONG DEFAULT 1, [M_EMAIL] TEXT(50), [M_FIRSTNAME] TEXT(100), [M_GLOW] TEXT(50) NULL, [M_GOLD] LONG DEFAULT 100, [M_GTOTAL] LONG DEFAULT 0, [M_HIDE_EMAIL] BYTE DEFAULT 0, [M_HOBBIES] MEMO, [M_HOMEPAGE] TEXT(50), [M_ICQ] TEXT(150), [M_IP] TEXT(50) DEFAULT '000.000.000.000', [M_KEY] TEXT(20), [M_LAST_IP] TEXT(255) DEFAULT '000.000.000.000', [M_LASTHEREDATE] TEXT(50), [M_LASTNAME] TEXT(100), [M_LASTPOSTDATE] TEXT(50), [M_LEVEL] INTEGER DEFAULT 1, [M_LINK1] TEXT(255), [M_LINK2] TEXT(255), [M_LNEWS] MEMO, [M_LOSSES] LONG DEFAULT 0, [M_MARSTATUS] TEXT(100), [M_MSN] TEXT(150), [M_NAME] TEXT(75), [M_NEWEMAIL] TEXT(50), [M_OCCUPATION] TEXT(255), [M_PAGE_VIEWS] LONG DEFAULT 0, [M_PASSWORD] TEXT(100), [M_PHOTO_URL] TEXT(255), [M_PMEMAIL] LONG DEFAULT 0, [M_PMRECEIVE] LONG DEFAULT 1, [M_POSTS] LONG DEFAULT 0, [M_QUOTE] MEMO, [M_RECEIVE_EMAIL] BYTE DEFAULT 1, [M_RECMAIL] INTEGER DEFAULT 0, [M_REP] LONG DEFAULT 5, [M_RNAME] TEXT(50), [M_RTOTAL] LONG DEFAULT 0, [M_SEX] TEXT(50), [M_SIG] MEMO, [M_STATE] TEXT(100), [M_STATUS] BYTE DEFAULT 1, [M_SUBSCRIPTION] BYTE DEFAULT 0, [M_TITLE] TEXT(50), [M_USERNAME] TEXT(150), [M_WINS] LONG DEFAULT 0, [M_YAHOO] TEXT(150), [M_ZIP] TEXT(20) NULL, [THEME_ID] TEXT(50) NULL, [M_SHOW_BIRTHDAY] LONG NULL DEFAULT 0, [M_PMSTATUS] INT DEFAULT 1, [M_PMBLACKLIST] MEMO, [M_DONATE] LONG DEFAULT 0, [M_LANG] TEXT(2), [M_LCID] LONG, [M_TIME_OFFSET] LONG DEFAULT 0, [M_TIME_TYPE] TEXT(2));"

createTable(checkIt(sSQL))

redim indexes(4)
indexes(0) = "CREATE INDEX [M_KEY] ON [" & strTablePrefix & "MEMBERS]([M_KEY]);"
indexes(1) = "CREATE INDEX [M_USERNAME] ON [" & strTablePrefix & "MEMBERS]([M_USERNAME]);"
indexes(2) = "CREATE INDEX [THEME_ID] ON [" & strTablePrefix & "MEMBERS]([THEME_ID]);"
indexes(3) = "CREATE INDEX [M_SHOW_BIRTHDAY] ON [" & strTablePrefix & "MEMBERS]([M_SHOW_BIRTHDAY]);"
indexes(4) = "CREATE INDEX [M_GLOW] ON [" & strTablePrefix & "MEMBERS]([M_GLOW]);"
createIndx(indexes)

	'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "MEMBERS "
		strSql = strSql & "(M_STATUS"
		strSql = strSql & ", M_NAME"
		strSql = strSql & ", M_USERNAME"
		strSql = strSql & ", M_PASSWORD"
		strSql = strSql & ", M_KEY"
		strSql = strSql & ", M_EMAIL"
		strSql = strSql & ", M_NEWEMAIL"
		strSql = strSql & ", M_COUNTRY" 
		strSql = strSql & ", M_HOMEPAGE"
		strSql = strSql & ", M_SIG"
		strSql = strSql & ", M_DEFAULT_VIEW" 
		strSql = strSql & ", M_LEVEL"
		strSql = strSql & ", M_AIM"
		strSql = strSql & ", M_YAHOO"
		strSql = strSql & ", M_MSN"
		strSql = strSql & ", M_ICQ"
		strSql = strSql & ", M_POSTS"
		strSql = strSql & ", M_DATE"
		strSql = strSql & ", M_LASTPOSTDATE"
		strSql = strSql & ", M_LASTHEREDATE"
		strSql = strSql & ", M_TITLE"
		strSql = strSql & ", M_SUBSCRIPTION"
		strSql = strSql & ", M_HIDE_EMAIL"
		strSql = strSql & ", M_RECEIVE_EMAIL"
		strSql = strSql & ", M_LAST_IP"
		strSql = strSql & ", M_IP"
		strSql = strSql & ", M_FIRSTNAME"
		strSql = strSql & ", M_LASTNAME"
		strSql = strSql & ", M_OCCUPATION"
		strSql = strSql & ", M_SEX"
		strSql = strSql & ", M_AGE"
		strSql = strSql & ", M_HOBBIES"
		strSql = strSql & ", M_LNEWS"
		strSql = strSql & ", M_QUOTE"
		strSql = strSql & ", M_BIO"
		strSql = strSql & ", M_MARSTATUS"
		strSql = strSql & ", M_LINK1"
		strSql = strSql & ", M_LINK2"
		strSql = strSql & ", M_CITY"
		strSql = strSql & ", M_PHOTO_URL"
		strSql = strSql & ", M_STATE"
		strSql = strSql & ", M_ZIP"
		strSql = strSql & ", M_PMEMAIL"
		strSql = strSql & ", M_PMRECEIVE"
		strSql = strSql & ", M_RECMAIL"
		strSql = strSql & ", M_GOLD"
		strSql = strSql & ", M_REP"
		strSql = strSql & ", M_LOSSES"
		strSql = strSql & ", M_WINS"
		strSql = strSql & ", M_GTOTAL"
		strSql = strSql & ", M_AVATAR_URL"
		strSql = strSql & ", M_RTOTAL"
		strSql = strSql & ", M_RNAME"
		strSql = strSql & ", M_PAGE_VIEWS"
		strSql = strSql & ", M_GLOW"
		strSql = strSql & ", THEME_ID"
		strSql = strSql & ", M_SHOW_BIRTHDAY"
		strSql = strSql & ", M_LCID"
		strSql = strSql & ", M_TIME_OFFSET"
		strSql = strSql & ", M_TIME_TYPE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1" 
		strSql = strSql & ", '" & adminName & "'"
		strSql = strSql & ", '" & adminName & "'"
		strSql = strSql & ", '" & adminPass & "'"  
		strSql = strSql & ", ' '"
		strSql = strSql & ", '" & adminEmail & "'" 'email
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", 'http://'"
		strSql = strSql & ", ''"
		strSql = strSql & ", 1"
		strSql = strSql & ", 3"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", 3"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", '" & DateToStr(now()) & "'"
		strSql = strSql & ", ''"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", '000.000.000.000'"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 'http://'"
		strSql = strSql & ", 'http://'"
		strSql = strSql & ", ''"
		strSql = strSql & ", 'http://'"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 100"
		strSql = strSql & ", 10"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/avatars/noavatar.gif'"
		strSql = strSql & ", 0"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'FF0000:FFFFFF'"
		strSql = strSql & ", '0'"
		strSql = strSql & ", 0"
		strSql = strSql & ", " & intPortalLCID
		strSql = strSql & ", " & timeoffset
		strSql = strSql & ", '12'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)

':::::::::::::::::::::::::::::::::: CREATE MEMBERS_PENDING TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "MEMBERS_PENDING")
sSQL = "CREATE TABLE [" & strTablePrefix & "MEMBERS_PENDING]([M_AGE] TEXT(10), [M_AIM] TEXT(150), [M_AVATAR_URL] TEXT(255), [M_BIO] MEMO, [M_CITY] TEXT(100), [M_COUNTRY] TEXT(40), [M_DATE] TEXT(50), [M_DEFAULT_VIEW] LONG DEFAULT 1, [M_EMAIL] TEXT(50), [M_FIRSTNAME] TEXT(100), [M_GLOW] TEXT(255) , [M_GOLD] LONG DEFAULT 100, [M_GTOTAL] LONG DEFAULT 0, [M_HIDE_EMAIL] BYTE DEFAULT 0, [M_HOBBIES] MEMO, [M_HOMEPAGE] TEXT(50), [M_ICQ] TEXT(150), [M_IP] TEXT(50) DEFAULT '000.000.000.000', [M_KEY] TEXT(20), [M_LAST_IP] TEXT(255) DEFAULT '000.000.000.000', [M_LASTHEREDATE] TEXT(50), [M_LASTNAME] TEXT(100), [M_LASTPOSTDATE] TEXT(50), [M_LEVEL] INTEGER DEFAULT 1, [M_LINK1] TEXT(255), [M_LINK2] TEXT(255), [M_LNEWS] MEMO, [M_LOSSES] LONG DEFAULT 0, [M_MARSTATUS] TEXT(100), [M_MSN] TEXT(150), [M_NAME] TEXT(75), [M_NEWEMAIL] TEXT(50), [M_OCCUPATION] TEXT(255), [M_PAGE_VIEWS] LONG DEFAULT 0, [M_PASSWORD] TEXT(100), [M_PHOTO_URL] TEXT(255), [M_PMEMAIL] LONG DEFAULT 0, [M_PMRECEIVE] LONG DEFAULT 1, [M_POSTS] LONG DEFAULT 0, [M_QUOTE] MEMO, [M_RECEIVE_EMAIL] BYTE DEFAULT 1, [M_RECMAIL] INTEGER DEFAULT 0, [M_REP] LONG DEFAULT 5, [M_RNAME] TEXT(50), [M_RTOTAL] LONG DEFAULT 0, [M_SEX] TEXT(50), [M_SIG] TEXT(255), [M_STATE] TEXT(100), [M_STATUS] BYTE DEFAULT 1, [M_SUBSCRIPTION] BYTE DEFAULT 0, [M_TITLE] TEXT(50), [M_USERNAME] TEXT(150), [M_WINS] LONG DEFAULT 0, [M_YAHOO] TEXT(150), [M_ZIP] TEXT(20), [THEME_ID] TEXT(50) NULL, [MEMBER_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [M_LCID] LONG, [M_TIME_OFFSET] LONG DEFAULT 0, [M_TIME_TYPE] TEXT(2));"

createTable(checkIt(sSQL))

redim indexes(0)
indexes(0) = "CREATE INDEX [M_KEY] ON [" & strTablePrefix & "MEMBERS_PENDING]([M_KEY]);"
createIndx(indexes)

':::::::::::::::::::::::: CREATE CP_CONFIG TABLE ::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "CP_CONFIG")
sSQL = "CREATE TABLE [" & strTablePrefix & "CP_CONFIG]([ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, [MAX_MY_TOPICS] LONG NOT NULL DEFAULT 5, [MEMBER_ID] LONG NOT NULL, [SHOW_MY_TOPICS] INT NOT NULL DEFAULT 1, [SHOW_PM] INT NOT NULL DEFAULT 1, [SHOW_RECENT_TOPICS] INT NOT NULL DEFAULT 1, [SHOW_STATUS] INT NOT NULL DEFAULT 1, [PM_OUTBOX] INT NOT NULL DEFAULT 1);"

createTable(checkIt(sSQL))

redim indexes(0)
indexes(0) = "CREATE INDEX [MEMBER_ID] ON [" & strTablePrefix & "CP_CONFIG]([MEMBER_ID]);"
createIndx(indexes)

'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "CP_CONFIG "
		strSql = strSql & "(MAX_MY_TOPICS"
		strSql = strSql & ", SHOW_RECENT_TOPICS"
		strSql = strSql & ", SHOW_MY_TOPICS"
		strSql = strSql & ", SHOW_PM"
		strSql = strSql & ", SHOW_STATUS"
		strSql = strSql & ", PM_OUTBOX"
		strSql = strSql & ", MEMBER_ID"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "5"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)

':::::::::::::::::::::::: CREATE PM TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "PM")
sSQL = "CREATE TABLE [" & strTablePrefix & "PM]([M_FROM] LONG, [M_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [M_MAIL] TEXT(50), [M_MESSAGE] MEMO, [M_OUTBOX] BYTE DEFAULT 1, [M_PMCOUNT] TEXT(50), [M_READ] LONG DEFAULT 0, [M_SENT] TEXT(50), [M_SUBJECT] TEXT(100), [M_TO] LONG, [M_SAVED] INT DEFAULT 0);"

createTable(checkIt(sSQL))
end sub 'tblMembers

sub tblCountries()
':::::::::::::::::::::: CREATE COUNTRIES TABLE :::::::::::::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "COUNTRIES")
sSQL="CREATE TABLE [" & strTablePrefix & "COUNTRIES]([CO_ABBREV] TEXT(255), [CO_CCTLD] TEXT(255), [CO_FLAG] TEXT(255), [CO_NAME] TEXT(255) PRIMARY KEY NOT NULL);"
'response.write checkIt(sSQL)
createTable(checkIt(sSQL))

'arrCntryData(1) = "[CO_NAME], [CO_ABBREV], [CO_CCTLD], [CO_FLAG]"
'-------------------- populate table with default values --------------------------
if isArray(arrCntryData) then
'Response.Write("Records: " & ubound(arrCntryData) & "<br />")
PopulateB(arrCntryData)
Response.Write(" 259 records added to table successfully<br />")
else
Response.Write("Not an array<br />")
end if
end sub 'tblCountries()

sub tblUploads()
':::::::::::::::::::::::: CREATE UPLOAD TABLE :::::::::::::::::::::::::::::::
droptable("" & strTablePrefix & "UPLOAD_CONFIG")
sSQL="CREATE TABLE [" & strTablePrefix & "UPLOAD_CONFIG]([UP_ACTIVE] TINYINT NOT NULL, [UP_ALLOWEDEXT] TEXT(255) NOT NULL, [UP_APPID] INTEGER NOT NULL, [UP_ALLOWEDUSERS] INTEGER NOT NULL, [UP_LOCATION] TEXT(255) NOT NULL, [UP_LOGFILE] TEXT(50) NOT NULL, [UP_LOGUSERS] TINYINT NOT NULL, [UP_SIZELIMIT] LONG NOT NULL, [ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL"
sSQL = sSQL & ", [UP_THUMB_MAX_W] LONG NOT NULL, [UP_THUMB_MAX_H] LONG NOT NULL, [UP_NORM_MAX_W] LONG NOT NULL, [UP_NORM_MAX_H] LONG NOT NULL, [UP_RESIZE] INT NOT NULL, [UP_CREATE_THUMB] INT NOT NULL, [UP_FOLDER] MEMO)"
'response.write checkIt(sSQL)
createTable(checkIt(sSQL))

		strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
		strSql = strSql & "(UP_SIZELIMIT"
		strSql = strSql & ", UP_ALLOWEDEXT"
		strSql = strSql & ", UP_LOGUSERS"
		strSql = strSql & ", UP_ALLOWEDUSERS"
		strSql = strSql & ", UP_LOCATION"
		strSql = strSql & ", UP_ACTIVE"
		strSql = strSql & ", UP_LOGFILE"
		strSql = strSql & ", UP_APPID"
		strSql = strSql & ", UP_THUMB_MAX_W"
		strSql = strSql & ", UP_THUMB_MAX_H"
		strSql = strSql & ", UP_NORM_MAX_W"
		strSql = strSql & ", UP_NORM_MAX_H"
		strSql = strSql & ", UP_RESIZE"
		strSql = strSql & ", UP_CREATE_THUMB"
		strSql = strSql & ", UP_FOLDER"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1000"
		strSql = strSql & ", 'gif,jpg,swf'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 4"
		strSql = strSql & ", 'banner'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 120"
		strSql = strSql & ", 120"
		strSql = strSql & ", 468"
		strSql = strSql & ", 60"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/banners/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)

		strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
		strSql = strSql & "(UP_SIZELIMIT"
		strSql = strSql & ", UP_ALLOWEDEXT"
		strSql = strSql & ", UP_LOGUSERS"
		strSql = strSql & ", UP_ALLOWEDUSERS"
		strSql = strSql & ", UP_LOCATION"
		strSql = strSql & ", UP_ACTIVE"
		strSql = strSql & ", UP_LOGFILE"
		strSql = strSql & ", UP_APPID"
		strSql = strSql & ", UP_THUMB_MAX_W"
		strSql = strSql & ", UP_THUMB_MAX_H"
		strSql = strSql & ", UP_NORM_MAX_W"
		strSql = strSql & ", UP_NORM_MAX_H"
		strSql = strSql & ", UP_RESIZE"
		strSql = strSql & ", UP_CREATE_THUMB"
		strSql = strSql & ", UP_FOLDER"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1000"
		strSql = strSql & ", 'gif,jpg'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'photo'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 200"
		strSql = strSql & ", 300"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/members/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		'populateA(strSql)

		strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
		strSql = strSql & "(UP_SIZELIMIT"
		strSql = strSql & ", UP_ALLOWEDEXT"
		strSql = strSql & ", UP_LOGUSERS"
		strSql = strSql & ", UP_ALLOWEDUSERS"
		strSql = strSql & ", UP_LOCATION"
		strSql = strSql & ", UP_ACTIVE"
		strSql = strSql & ", UP_LOGFILE"
		strSql = strSql & ", UP_APPID"
		strSql = strSql & ", UP_THUMB_MAX_W"
		strSql = strSql & ", UP_THUMB_MAX_H"
		strSql = strSql & ", UP_NORM_MAX_W"
		strSql = strSql & ", UP_NORM_MAX_H"
		strSql = strSql & ", UP_RESIZE"
		strSql = strSql & ", UP_CREATE_THUMB"
		strSql = strSql & ", UP_FOLDER"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1000"
		strSql = strSql & ", 'gif,jpg'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'avatar'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'avatar.txt'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 64"
		strSql = strSql & ", 64"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/members/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		'populateA(strSql)		

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
end sub 'tblUploads()

sub tblAPPS()	
':::::::::::::::: CREATE " & strTablePrefix & "APPS TABLE ::::::::::::::::::::::::::
droptable("" & strTablePrefix & "APPS")
sSQL = "CREATE TABLE [" & strTablePrefix & "APPS]([APP_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [APP_NAME] TEXT(100), [APP_INAME] TEXT(100) NULL, [APP_DEBUG] INT, [APP_CONFIG] TEXT(100), [APP_ACTIVE] INT, [APP_GROUPS_USERS] MEMO NULL, [APP_SUBSCRIPTIONS] INT, [APP_BOOKMARKS] INT, [APP_VIEW] MEMO NULL, [APP_UFOLDER] MEMO NULL"
' extra app config integer fields
sSQL = sSQL & ", [APP_iDATA1] INT, [APP_iDATA2] INT, [APP_iDATA3] INT, [APP_iDATA4] INT, [APP_iDATA5] INT, [APP_iDATA6] INT, [APP_iDATA7] INT, [APP_iDATA8] INT, [APP_iDATA9] INT, [APP_iDATA10] INT"
' extra app config memo fields
sSQL = sSQL & ", [APP_tDATA1] MEMO NULL, [APP_tDATA2] MEMO NULL, [APP_tDATA3] MEMO NULL, [APP_tDATA4] MEMO NULL, [APP_tDATA5] MEMO NULL);"
createTable(checkIt(sSQL))

redim arrData(2)
arrData(0) = "[" & strTablePrefix & "APPS]"
arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_iDATA1],[APP_iDATA2],[APP_iDATA3],[APP_iDATA4],[APP_iDATA5],[APP_iDATA6]"
arrData(2) = "'PM','PM',1,0,'1,2',3,3,'config_pm',0,30,50,0,0,1"
populateB(arrData)

createIndex("CREATE INDEX [APP_INAME] ON [" & strTablePrefix & "APPS]([APP_INAME]);")
'createIndex("CREATE INDEX [APP_GROUPS] ON [" & strTablePrefix & "APPS]([APP_GROUPS]);")
	'response.Write("<hr /><hr />")
end sub

sub portalFrontPage()	
':::::::::::::::: CREATE " & strTablePrefix & "FP TABLE ::::::::::::::::::::::::::
droptable("" & strTablePrefix & "FP")
sSQL = "CREATE TABLE [" & strTablePrefix & "FP]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [FP_NAME] TEXT(100) NULL, [FP_INAME] TEXT(100) NULL, [FP_FUNCTION] TEXT(50) NULL, [FP_ACTIVE] INT, [FP_DESC] MEMO NULL, [FP_COLUMN] INT, [FP_GROUPS] MEMO NULL, [APP_ID] INT, [FP_STICKY] INT);"
createTable(checkIt(sSQL))

'createIndex("CREATE INDEX [G_ID] ON [" & strTablePrefix & "GROUPS]([G_ID]);")

'response.Write("<h4>Portal Front Page default data</h4><br />")
redim arrData(11)
arrData(0) = "[" & strTablePrefix & "FP]"
arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
arrData(2) = "'" & txtFPWelMsg & "','welcome','welcome_fp',1,2,'" & txtFPWelMsgDesc & "','1,2,3',0"
arrData(3) = "'" & txtFPAnnounce & "','announcements','announce_fp',1,2,'" & txtFPAnnounceDesc & "','1,2,3',0"
arrData(4) = "'" & txtFPMainMnu & "','main_menu','menu_fp',1,4,'" & txtFPMainMnuDesc & "','1,2,3',0"
arrData(5) = "'" & txtFPSkinSel & "','theme_changer','theme_changer',1,4,'" & txtFPSkinSelDesc & "','3',0"
arrData(6) = "'" & txtFPAff & "','affiliates','affiliateBanners',1,4,'" & txtFPAffDesc & "','1,2,3',0"
arrData(7) = "'Support SkyPortal','support_skyportal','others_fp',1,4,'" & txtFPOther & "','1,2,3',0"
arrData(8) = "'" & txtFPProj & "','projects','projects_fp',1,4,'" & txtFPProjDesc & "','1,2,3',0"
arrData(9) = "'" & txtFPSrch & "','site_search','search_fp',1,4,'" & txtFPSrchDesc & "','1,2,3',0"
arrData(10) = "'" & txtFPLoginBlk & "','login_box','login_box',0,4,'" & txtFPLoginBlkDesc & "','3',0"
arrData(11) = "'Rate SkyPortal','aspin','m_aspin',1,4,'" & txtFPLoginBlkDesc & "','1,2,3',0"

populateB(arrData)
	'response.Write("<hr /><hr />")
end sub

':::::::::::::::: CREATE " & strTablePrefix & "FP_USERS TABLE ::::::::::::::::::::::::::
sub portalFPusers()	
droptable("" & strTablePrefix & "FP_USERS")
sSQL = "CREATE TABLE [" & strTablePrefix & "FP_USERS]([uid] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [fp_uid] INT, [fp_leftcol] MEMO NULL, [fp_maincol] MEMO NULL, [fp_rightcol] MEMO NULL, [fp_leftsticky] MEMO NULL, [fp_mainsticky] MEMO NULL, [fp_rightsticky] MEMO NULL);"
createTable(checkIt(sSQL))

createIndex("CREATE INDEX [fp_uid] ON [" & strTablePrefix & "FP_USERS]([fp_uid]);")
	redim arrData(2)
	arrData(0) = "[" & strTablePrefix & "FP_USERS]"
	arrData(1) = "[fp_uid],[fp_leftcol],[fp_rightcol],[fp_leftsticky],[fp_rightsticky],[fp_mainsticky],[fp_maincol]"
	arrData(2) = "0,'" & txtFPSkinSel & ":theme_changer','" & txtFPSrch & ":search_fp','" & txtFPMainMnu & ":menu_fp','Support SkyPortal:others_fp,Rate SkyPortal:m_aspin," & txtFPAff & ":affiliateBanners','" & txtFPAnnounce & ":announce_fp','" & txtFPWelMsg & ":welcome_fp'"
	populateB(arrData)
	'response.Write("<hr /><hr />")
end sub

':::::::::::::::: CREATE " & strTablePrefix & "GROUPS TABLE ::::::::::::::::::::::::::
sub portalGroups()	
droptable("" & strTablePrefix & "GROUPS")
sSQL = "CREATE TABLE [" & strTablePrefix & "GROUPS]([G_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [G_NAME] TEXT(100), [G_INAME] TEXT(100) NULL, [G_DESC] MEMO, [G_CREATE] TEXT(100), [G_MODIFIED] TEXT(100), [G_ACTIVE] INT, [G_ADDMEM] INT);"
createTable(checkIt(sSQL))

'createIndex("CREATE INDEX [G_ID] ON [" & strTablePrefix & "GROUPS]([G_ID]);")

redim arrData(5)
arrData(0) = "[" & strTablePrefix & "GROUPS]"
arrData(1) = "[G_NAME],[G_INAME],[G_DESC],[G_CREATE],[G_ACTIVE],[G_ADDMEM]"
arrData(2) = "'" & txtSUAdminist & "','Administrator','" & txtSUAdministDesc & "','" & DateToStr(now()) & "',0,1"
arrData(3) = "'" & txtSUMember & "','Members','" & txtSUMemberDesc & "','" & DateToStr(now()) & "',0,0"
arrData(4) = "'" & txtSUEvOne & "','Guests','" & txtSUEvOneDesc & "','" & DateToStr(now()) & "',0,0"
arrData(5) = "'" & txtSUModerator & "','Moderator','" & txtSUModDesc & "','" & DateToStr(now()) & "',0,1"
populateB(arrData)
	'response.Write("<hr /><hr />")
end sub

':::::::::::::::: CREATE " & strTablePrefix & "GROUP_MEMBERS TABLE ::::::::::::::::::::::::::
sub portalGroupMembers()	
droptable("" & strTablePrefix & "GROUP_MEMBERS")
sSQL = "CREATE TABLE [" & strTablePrefix & "GROUP_MEMBERS]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [G_MEMBER_ID] int, [G_GROUP_ID] int, [G_GROUP_LEADER] int);"
createTable(checkIt(sSQL))

createIndex("CREATE INDEX [G_MEMBER_ID] ON [" & strTablePrefix & "GROUP_MEMBERS]([G_MEMBER_ID]);")
createIndex("CREATE INDEX [G_GROUP_ID] ON [" & strTablePrefix & "GROUP_MEMBERS]([G_GROUP_ID]);")

'lets populate the administrator group with the site superadmin(s) and admin(s)
' get admin id's from the db
  sSQL = "SELECT MEMBER_ID, M_NAME FROM " & strTablePrefix & "MEMBERS WHERE M_LEVEL = 3 AND M_STATUS = 1"
  set rsAdmin = my_Conn.execute(sSQL)
  if not rsAdmin.eof then
    do until rsAdmin.eof
      rID = rsAdmin("MEMBER_ID")
      rName = lcase(rsAdmin("M_NAME")) & ","
	  rSA = 0
	
	  if instr(strWebMaster,rName) > 0 then
	    rSA = 1
	  end if
	
      strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS "
	  strSql = strSql & "(G_MEMBER_ID,G_GROUP_ID,G_GROUP_LEADER) VALUES "
	  strSql = strSql & "(" & rID & ",1," & rSA & ");"
	  executeThis(strSql)
	  rsAdmin.movenext
	loop
  end if
  set rsAdmin = nothing

'lets populate the MODERATOR group with the site forum moderators(s)
' get the moderator group ID number
  sSQL = "SELECT G_ID FROM " & strTablePrefix & "GROUPS WHERE G_INAME = 'Moderator'"
  set rsGID = my_Conn.execute(sSQL)
    intGID = rsGID(0)
  set rsGID = nothing

' get moderator id's from the db
  sSQL = "SELECT Member_ID FROM " & strTablePrefix & "MEMBERS WHERE M_LEVEL = 2 AND M_STATUS = 1"
  set rsAdmin = my_Conn.execute(sSQL)
  if not rsAdmin.eof then
    do until rsAdmin.eof
      rID = rsAdmin("Member_ID")
	
      strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS "
	  strSql = strSql & "(G_MEMBER_ID,G_GROUP_ID,G_GROUP_LEADER) VALUES "
	  strSql = strSql & "(" & rID & "," & intGID & ",0);"
	  executeThis(strSql)
	  rsAdmin.movenext
	loop
  end if
  set rsAdmin = nothing

redim arrData(2)
arrData(0) = "[" & strTablePrefix & "GROUP_MEMBERS]"
arrData(1) = "[G_MEMBER_ID],[G_GROUP_ID],[G_GROUP_LEADER]"
arrData(2) = "1,1,0"
'populateB(arrData)
end sub

sub portalGroupPerms()
'droptable("" & strTablePrefix & "GROUP_PERMS")
sSQL = "CREATE TABLE [" & strTablePrefix & "GROUP_PERMS]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [APP_ID] LONG NOT NULL, [G_ID] LONG NOT NULL, [APP_CAT] LONG NULL DEFAULT 0, [APP_SUB] LONG NULL DEFAULT 0, [G_LEADER] LONG NULL DEFAULT 0, [GP_READ] LONG NULL DEFAULT 1, [GP_WRITE] LONG NULL DEFAULT 0, [GP_DELETE] LONG NULL DEFAULT 0, [GP_MODIFY] LONG NULL DEFAULT 0);"

'createTable(checkIt(sSQL))
  
	redim arrData(3)
	arrData(0) = "[" & strTablePrefix & "GROUP_PERMS]"
	arrData(1) = "[APP_ID],[G_ID],[APP_CAT],[APP_SUB],[G_LEADER],[GP_READ],[GP_WRITE],[GP_DELETE],[GP_MODIFY]"
	arrData(2) = "1,1,0,0,0,1,1,1,1"
	arrData(3) = "1,2,0,0,0,1,1,1,1"
	'populateB(arrData)
end sub

sub sky_Pages()
'::::::::::::::::::: CREATE PORTAL_PAGES  TABLE :::::::::::::::::::::::::
   response.Write("<hr /><h5>Custom_Pages default data</h5>")
droptable("" & strTablePrefix & "PAGES")
sSQL = "CREATE TABLE [" & strTablePrefix & "PAGES]([P_ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, [P_NAME] TEXT(255), [P_INAME] TEXT(255), [P_TITLE] TEXT(255), [P_CONTENT] MEMO NULL, [P_ACONTENT] MEMO NULL, [P_LEFTCOL] MEMO NULL, [P_RIGHTCOL] MEMO NULL, [P_MAINTOP] MEMO NULL, [P_MAINBOTTOM] MEMO NULL, [P_APP] INT, [P_USE_PG_DISP] INT, [P_OTHER_URL] TEXT(255) NULL, [P_CAN_DELETE] INT, [P_META_TITLE] TEXT(255) NULL, [P_META_DESC] TEXT(255) NULL, [P_META_KEY] TEXT(255) NULL, [P_META_EXPIRES] TEXT(255) NULL, [P_META_RATING] TEXT(255) NULL, [P_META_DIST] TEXT(255) NULL, [P_META_ROBOTS] TEXT(255) NULL);"
createTable(checkIt(sSQL))
		
		'P_META_TITLE
		'P_META_DESC
		'P_META_KEY
		'P_META_EXPIRES
		'P_META_RATING
		'P_META_DIST
		'P_META_ROBOTS

'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strTablePrefix & "PAGES "
		strSql = strSql & "(P_NAME"
		strSql = strSql & ", P_INAME"
		strSql = strSql & ", P_TITLE"
		strSql = strSql & ", P_CONTENT"
		strSql = strSql & ", P_LEFTCOL"
		strSql = strSql & ", P_RIGHTCOL"
		strSql = strSql & ", P_MAINTOP"
		strSql = strSql & ", P_MAINBOTTOM"
		strSql = strSql & ", P_APP"
		strSql = strSql & ", P_USE_PG_DISP"
		strSql = strSql & ", P_OTHER_URL"
		strSql = strSql & ", P_CAN_DELETE"
		strSql = strSql & ", P_META_TITLE"
		strSql = strSql & ", P_META_DESC"
		strSql = strSql & ", P_META_KEY"
		strSql = strSql & ", P_META_EXPIRES"
		strSql = strSql & ", P_META_RATING"
		strSql = strSql & ", P_META_DIST"
		strSql = strSql & ", P_META_ROBOTS"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txtSitePol & "'"
		strSql = strSql & ", 'policy'"
		strSql = strSql & ", '" & txtSitePol & "'"
		strSql = strSql & ", '" & txtPolicyHTML & "'"
		strSql = strSql & ", 'Main Menu:menu_fp'"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'policy.asp'"
		strSql = strSql & ", 0"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO " & strTablePrefix & "PAGES "
		strSql = strSql & "(P_NAME"
		strSql = strSql & ", P_INAME"
		strSql = strSql & ", P_TITLE"
		strSql = strSql & ", P_CONTENT"
		strSql = strSql & ", P_LEFTCOL"
		strSql = strSql & ", P_RIGHTCOL"
		strSql = strSql & ", P_MAINTOP"
		strSql = strSql & ", P_MAINBOTTOM"
		strSql = strSql & ", P_APP"
		strSql = strSql & ", P_USE_PG_DISP"
		strSql = strSql & ", P_OTHER_URL"
		strSql = strSql & ", P_CAN_DELETE"
		strSql = strSql & ", P_META_TITLE"
		strSql = strSql & ", P_META_DESC"
		strSql = strSql & ", P_META_KEY"
		strSql = strSql & ", P_META_EXPIRES"
		strSql = strSql & ", P_META_RATING"
		strSql = strSql & ", P_META_DIST"
		strSql = strSql & ", P_META_ROBOTS"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txtPrivacySt & "'"
		strSql = strSql & ", 'privacy'"
		strSql = strSql & ", '" & txtPrivacySt & "'"
		strSql = strSql & ", '" & txtPrivacyHTML & "'"
		strSql = strSql & ", 'Main Menu:menu_fp'"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'privacy.asp'"
		strSql = strSql & ", 0"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"
		strSql = strSql & ", ''"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub sky_menu()
'response.Write("<a href=""admin_menu.asp"">Admin Menu</a>")

'droptable("Menu")
	
	if bFso then
	 on error resume next
	 Err.Clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   rt = server.mappath("files")
	   fso.CreateFolder(rt & "/config")
	   if fso.FolderExists(rt & "/config") then
	     fso.CreateFolder(rt & "/config/menu")
	   end if
	   set fso = nothing
	 end if
	 Err.Clear
	 on error goto 0
	end if
	
sSql = "CREATE TABLE [Menu]([id] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [app_id] LONG DEFAULT 0, [INAME] TEXT(255), [Link] TEXT(255), [mnuAccess] MEMO, [mnuAdd] MEMO, [mnuFunction] TEXT(255), [mnuImage] TEXT(255), [mnuOrder] LONG DEFAULT 1, [mnuTitle] TEXT(255), [Name] TEXT(255), [onclick] TEXT(255), [Parent] TEXT(255), [ParentID] LONG DEFAULT 0, [Target] TEXT(255));"
createTable(checkIt(sSQL))

admin_menu()
response.Write("<br />b_members<br />")
b_members()
response.Write("<br />nav_menu<br />")
nav_menu()
response.Write("<br />main_menu<br />")
main_menu()
response.Write("<br />cp_menu<br />")
cp_menu()

mnu.DelMenuFiles("")
end sub

sub cp_menu()
  ':: start button template
  mnuName = txtMnuCP	
  mnuINAME = "cp_main"
  mnu_icon = "Themes/<%= strTheme %" & ">/icons/arrow1.gif"
  
  redim arrData(12)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuFunction,mnuTitle,INAME,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtHome & "', '"& mnuINAME &"','default.asp','Themes/<%= strTheme %" & ">/icons/arrow1.gif','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',1"
  arrData(3) = "'" & txtPersSet & "', '"& mnuINAME &"','cp_main.asp?cmd=5','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',2"
  arrData(4) = "'" & txtPsnlMsgs & "', '"& mnuINAME &"','pm.asp','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',3"
  arrData(5) = "'" & txtEditProf & "', '"& mnuINAME &"','cp_main.asp?cmd=9','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',4"
  arrData(6) = "'" & txtViewProf & "', '"& mnuINAME &"','cp_main.asp','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',5"
  arrData(7) = "'" & txtEditAvatar & "', '"& mnuINAME &"','cp_main.asp?cmd=1','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',6"
  arrData(8) = "'" & txtMyBkmks & "', '"& mnuINAME &"','cp_main.asp?cmd=7','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',7"
  arrData(9) = "'" & txtMySubsc & "', '"& mnuINAME &"','cp_main.asp?cmd=6','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',8"
  arrData(10) = "'" & txtMyRecTop & "', '"& mnuINAME &"','cp_main.asp?cmd=4','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',9"
  arrData(11) = "'" & txtAdminOpts & "', '"& mnuINAME &"','admin_home.asp','" & mnu_icon & "','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'1',10"
  arrData(12) = "'" & txtPndTsks & "', '"& mnuINAME &"','admin_home.asp','" & mnu_icon & "','','_parent','cntPendTsks()','"& mnuName &"','"& mnuINAME &"',0,'1',11"
  populateB(arrData)
end sub

sub main_menu()
  ':: start button template
  mnuName = txtMnuDefault	
  mnuINAME = "def_main"

  redim arrData(9)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuFunction,mnuTitle,INAME,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtRegister & "', '"& mnuINAME &"','policy.asp','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'3',2"
  arrData(3) = "'" & txtHome & "', '"& mnuINAME &"','default.asp','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'',1"
  arrData(4) = "'" & txtMsgs & " ', '"& mnuINAME &"','pm.asp','','','_parent','newPM','"& mnuName &"','"& mnuINAME &"',1,'1,2',3"
  arrData(5) = "'" & txtMyBkmks & "', '"& mnuINAME &"','cp_main.asp?cmd=7','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'1,2',4"
  arrData(6) = "'" & txtMySubsc & "', '"& mnuINAME &"','cp_main.asp?cmd=6','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'1,2',5"
  arrData(7) = "'" & txtSStats & "', '"& mnuINAME &"','statistics.asp','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'1',6"
  arrData(8) = "'" & txtAdminOpts & "', '"& mnuINAME &"','admin_home.asp','','','_parent','','"& mnuName &"','"& mnuINAME &"',0,'1',7"
  arrData(9) = "'" & txtPndTsks & "', '"& mnuINAME &"','admin_home.asp','','','_parent','cntPendTsks()','"& mnuName &"','"& mnuINAME &"',0,'1',8"
  
  populateB(arrData)
end sub

sub nav_menu()
  mnuName = txtMnuNav	
  mnuINAME = "nav_main"

  ':: HOME Button
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtHome & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & txtHome & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(4)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtHome & "', '" & txtHome & "','default.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtRegister & "', '" & txtHome & "','policy.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtContactUs & "', '" & txtHome & "','','_parent','openWindowPM(''pm_pop.asp'')','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',3"
  populateB(arrData)
 
 ':: add MEMBER Button REFERENCE
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAdd,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuMbr & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','b_members','',2"
  populateB(arrData)

end sub

sub b_ipgate()
  ':: start button template
  mnuName = "* " & txtMnuIPGateAdm	
  mnuINAME = "b_ipgate"
  mnuBName = txtMnuIpGate

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(11)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuIpMain & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=MainMenu','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuIpAdmin & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=adminip','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuIpUsrBan & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=UserSettings','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',3"
  arrData(5) = "'" & txtMnuIpBan & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=IPBanning','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',4"
  arrData(6) = "'" & txtMnuIpVwLgs & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=Logs','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',5"
  arrData(7) = "'" & txtMnuIpArchLogs & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=logarchive','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',6"
  arrData(8) = "'" & txtMnuIpErOldLogs & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=deletelog&qry=15','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',7"
  arrData(9) = "'" & txtMnuIpEdBlkPgs & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=pagekeys','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',8"
  arrData(10) = "'" & txtMnuIpSetngs & "', '" & mnuBName & "','admin_ipgate.asp?ViewPage=Settings','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',9"
  arrData(11) = "'" & txtMnuIpHelp & "', '" & mnuBName & "','','_parent','openWindow5(''pop_help.asp?mode=3'');','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',10"
  populateB(arrData)
end sub

sub b_avatar()
  ':: start button template
  mnuName = "* " & txtMnuAvAdmin	
  mnuINAME = "b_avatar_cfg"
  mnuBName = txtMnuAvSetup

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(6)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuAvSetngs & "', '" & mnuBName & "','admin_avatar_home.asp?mode=avset','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuAvAdd & "', '" & mnuBName & "','admin_avatar_home.asp?mode=added','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuAvSync & "', '" & mnuBName & "','admin_avatar_home.asp?mode=avsync','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',3"
  arrData(5) = "'" & txtMnuAvRevEd & "', '" & mnuBName & "','admin_avatar_home.asp?mode=avrev','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',4"
  arrData(6) = "'" & txtMnuAvUpld & "', '" & mnuBName & "','admin_avatar_home.asp?mode=avupld','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',5"
  populateB(arrData)
end sub

sub b_flags()
  ':: start button template
  mnuName = "* " & txtMnuCf	
  mnuINAME = "b_flags"
  mnuBName = txtMnuCfCntrs

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuCfAll & "', '" & mnuBName & "','admin_countries.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuCfAdd & "', '" & mnuBName & "','','_parent','show(''ab'');hide(''aa'');','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)
end sub

sub b_banner_cfg()
  ':: start button template
  mnuName = "* " & txtMnuBnAdmin	
  mnuINAME = "b_banners"
  mnuBName = txtMnuBn

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(5)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuBnView & "', '" & mnuBName & "','admin_banner_manager.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuBnAdd & "', '" & mnuBName & "','admin_banner_manager.asp?mode=7','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuBnAfView & "', '" & mnuBName & "','admin_banner_manager.asp?loc=2','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',3"
  arrData(5) = "'" & txtMnuBnAfAdd & "', '" & mnuBName & "','admin_banner_manager.asp?mode=7&loc=2','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',4"
  populateB(arrData)
end sub

sub b_pm()
  ':: start button template
  mnuName = "* " & txtMnuPmAdmin	
  mnuINAME = "b_pm_cfg"
  mnuBName = txtMnuPmMgr

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuPmCfg & "', '" & mnuBName & "','','_parent','show(''pbb'');show(''paa'');hide(''pcc'');','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuPmNewUsrs & "', '" & mnuBName & "','','_parent','show(''pcc'');hide(''paa'');hide(''pbb'');','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)
end sub

sub b_layout_mgr()
  ':: start button template
  mnuName = "* " & txtMnuCpAdmin	
  mnuINAME = "b_layout"
  mnuBName = txtMnuCpMgr

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(7)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuCpHP & "', '" & mnuBName & "','admin_config_fp.asp?cmd=3','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuCpCusPg & "', '" & mnuBName & "','admin_config_cp.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuCpActBlk & "', '" & mnuBName & "','admin_config_fp.asp?cmd=1','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',3"
  arrData(5) = "'" & txtMnuCpInActBlk & "', '" & mnuBName & "','admin_config_fp.asp?cmd=0','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',4"
  arrData(6) = "'" & txtMnuCpAddNewBlk & "', '" & mnuBName & "','admin_config_fp.asp?cmd=2','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',5"
  arrData(7) = "'" & txtMnuCpResetMbrs & "', '" & mnuBName & "','admin_config_fp.asp?cmd=&mode=5','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',6"
  populateB(arrData)
end sub

sub b_members()
  ':: start MEMBERS button
  mnuName = "* " & txtMnuMMbr
  mnuINAME = "b_members"
  mnuBName = txtMnuMMbr

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','1,2',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(12)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuMMax & "', '" & mnuBName & "','','_parent','mymax()','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',1"
  arrData(3) = "'" & txtMnuMCp & "', '" & mnuBName & "','cp_main.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',2"
  arrData(4) = "'" & txtMnuMMyProf & "', '" & mnuBName & "','','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',3"
  arrData(5) = "'" & txtMnuMMsgs & "', '" & mnuBName & "','','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",1,'',4"
  arrData(6) = "'" & txtMnuMBkmk & "', '" & mnuBName & "','cp_main.asp?cmd=7','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',5"
  arrData(7) = "'" & txtMnuMSubsc & "', '" & mnuBName & "','cp_main.asp?cmd=6','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',6"
  arrData(8) = "'" & txtMnuMMbrLst & "', '" & mnuBName & "','members.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',7"
  arrData(9) = "'" & txtMnuMActUsrs & "', '" & mnuBName & "','active_users.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',8"
  arrData(10) = "'" & txtMnuMSiteMntr & "', '" & mnuBName & "','site_monitor.asp','_search','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',9"
  arrData(11) = "'" & txtMnuMAdmnOpts & "', '" & mnuBName & "','admin_home.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',10"
  arrData(12) = "'" & txtMnuMRptPst & "', '" & mnuBName & "','forum_report_post_moderate.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',11"
  populateB(arrData)

  ':: 'My Profile' sublinks
  sSql = "select ID from menu where Name = '" & txtMnuMMyProf & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(4)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,  mnuOrder"
  arrData(2) = "'" & txtMnuMViewProf & "', '" & txtMnuMMyProf & "','cp_main.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuMEdProf & "', '" & txtMnuMMyProf & "','cp_main.asp?cmd=9','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuMEdAv & "', '" & txtMnuMMyProf & "','cp_main.asp?cmd=1&mode=AvatarEdit','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)

  ':: 'Messages' sublinks
  sSql = "select ID from menu where Name = '" & txtMnuMMsgs & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,  mnuOrder"
  arrData(2) = "'" & txtMnuMViewInbx & "', '" & txtMnuMMsgs & "','pm.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuMCpos & "', '" & txtMnuMMsgs & "','pm.asp?cmd=2','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)
  
':: end MEMBERS button
end sub

sub b_members2()
  ':: start MEMBERS button
  mnuName = "* " & txtMnuMMbr
  mnuINAME = "b_members"
  mnuBName = txtMnuMMbr

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','1,2',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(12)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'" & txtMnuMMax & "', '" & mnuBName & "','','_parent','mymax()','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',1"
  arrData(3) = "'" & txtMnuMCp & "', '" & mnuBName & "','cp_main.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',2"
  arrData(4) = "'" & txtMnuMMyProf & "', '" & mnuBName & "','','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',3"
  arrData(5) = "'" & txtMnuMMsgs & "', '" & mnuBName & "','','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",1,'',4"
  arrData(6) = "'" & txtMnuMBkmk & "', '" & mnuBName & "','cp_main.asp?cmd=7','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',5"
  arrData(7) = "'" & txtMnuMSubsc & "', '" & mnuBName & "','cp_main.asp?cmd=6','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',6"
  arrData(8) = "'" & txtMnuMMbrLst & "', '" & mnuBName & "','members.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',7"
  arrData(9) = "'" & txtMnuMActUsrs & "', '" & mnuBName & "','active_users.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'',8"
  arrData(10) = "'" & txtMnuMSiteMntr & "', '" & mnuBName & "','site_monitor.asp','_search','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',9"
  arrData(11) = "'" & txtMnuMAdmnOpts & "', '" & mnuBName & "','admin_home.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',10"
  arrData(12) = "'" & txtMnuMRptPst & "', '" & mnuBName & "','forum_report_post_moderate.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",0,'1',11"
  populateB(arrData)

  ':: 'My Profile' sublinks
  sSql = "select ID from menu where Name = '" & txtMnuMMyProf & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(4)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,  mnuOrder"
  arrData(2) = "'" & txtMnuMViewProf & "', '" & txtMnuMMyProf & "','cp_main.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuMEdProf & "', '" & txtMnuMMyProf & "','cp_main.asp?cmd=9','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  arrData(4) = "'" & txtMnuMEdAv & "', '" & txtMnuMMyProf & "','cp_main.asp?cmd=1&mode=AvatarEdit','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)

  ':: 'Messages' sublinks
  sSql = "select ID from menu where Name = '" & txtMnuMMsgs & "' and INAME = '" & mnuINAME & "'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuAccess,  mnuOrder"
  arrData(2) = "'" & txtMnuMViewInbx & "', '" & txtMnuMMsgs & "','pm.asp','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',1"
  arrData(3) = "'" & txtMnuMCpos & "', '" & txtMnuMMsgs & "','pm.asp?cmd=2','_parent','','','" & mnuName & "','" & mnuINAME & "'," & pID & ",'',2"
  populateB(arrData)
  
':: end MEMBERS button
end sub

sub b_site_config()
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name, Parent, Link, mnuImage, onclick, Target, mnuTitle, INAME, mnuOrder"
arrData(2) = "'" & txtMnuSASiteCfg & "', 'b_site_cfg','','','','','* " & txtMnuSACfgAdmin & "','b_site_cfg',1"
populateB(arrData)

sSql = "select ID from menu where Name = '" & txtMnuSASiteCfg & "' and INAME = 'b_site_cfg'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(10)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuOrder"
arrData(2) = "'" & txtMnuSAAdHm & "', '" & txtMnuSASiteCfg & "','admin_home.asp','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",1"
arrData(3) = "'" & txtMnuSAGenSet & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=1','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",2"
arrData(4) = "'" & txtMnuSABWFltr & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=2','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",3"
arrData(5) = "'" & txtMnuSADtTm & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=3','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",4"
arrData(6) = "'" & txtMnuSANtFeat & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=9','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",5"
arrData(7) = "'" & txtMnuSAEmlSrvr & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=4','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",6"
arrData(8) = "'" & txtMnuSAEmlMbrs & "', '" & txtMnuSASiteCfg & "','admin_emaillist.asp','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",7"
arrData(9) = "'" & txtMnuSASrvrInfo & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=7','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",8"
arrData(10) = "'" & txtMnuSASiteVars & "', '" & txtMnuSASiteCfg & "','admin_home.asp?cmd=8','_parent','','','* " & txtMnuSACfgAdmin & "','b_site_cfg'," & pID & ",9"
populateB(arrData)
':: end site configuration button
end sub

sub b_mem_config()
':: start MEMBERS button
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuOrder"
arrData(2) = "'" & txtMnuMMbr & "', 'b_mem_cfg','','','','','* " & txtMnuSAMbrAdmin & "','b_mem_cfg',2"
populateB(arrData)

sSql = "select ID from menu where Name = '" & txtMnuMMbr & "' and INAME = 'b_mem_cfg'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(5)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,onclick,mnuImage,mnuTitle,INAME,ParentID,mnuOrder"
arrData(2) = "'" & txtMnuSAMDet & "', '" & txtMnuMMbr & "','admin_config_members.asp','_parent','','','* " & txtMnuSAMbrAdmin & "','b_mem_cfg'," & pID & ",1"
arrData(3) = "'" & txtMnuSAMbrRnk & "', '" & txtMnuMMbr & "','admin_config_members.asp?cmd=1','_parent','','','* " & txtMnuSAMbrAdmin & "','b_mem_cfg'," & pID & ",2"
arrData(4) = "'" & txtMnuSAMbrPend & "', '" & txtMnuMMbr & "','admin_accounts_pending.asp','_parent','','','* " & txtMnuSAMbrAdmin & "','b_mem_cfg'," & pID & ",3"
arrData(5) = "'" & txtMnuSAMbrClean & "', '" & txtMnuMMbr & "','admin_config_members.asp?cmd=2','_parent','','','* " & txtMnuSAMbrAdmin & "','b_mem_cfg'," & pID & ",4"
populateB(arrData)
':: end MEMBERS button
end sub

sub b_manager_config()
':: start MANAGERS button
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuOrder"
arrData(2) = "'" & txtMnuSAMgr & "', 'b_managers','','','','','* " & txtMnuSAMgrAdmin & "','b_managers',3"
populateB(arrData)

sSql = "select ID from menu where Name = '" & txtMnuSAMgr & "' and INAME = 'b_managers'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(15)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,mnuImage,onclick,mnuTitle,INAME,ParentID,mnuOrder"
arrData(2) = "'" & txtMnuSAMgrLayout & "', '" & txtMnuSAMgr & "','admin_config_cp.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",1"
arrData(3) = "'" & txtMnuSAMgrModule & "', '" & txtMnuSAMgr & "','admin_config_modules.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",2"
arrData(4) = "'" & txtMnuSAMgrGrp & "', '" & txtMnuSAMgr & "','admin_config_groups.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",3"
arrData(5) = "'" & txtMnuSAMgrMenu & "', '" & txtMnuSAMgr & "','admin_menu.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",4"
arrData(6) = "'" & txtMnuSAMgrBanr & "', '" & txtMnuSAMgr & "','admin_banner_manager.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",5"
arrData(7) = "'" & txtMnuSAMgrSkins & "', '" & txtMnuSAMgr & "','admin_skins_config.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",6"
arrData(8) = "'" & txtMnuSAMgrUploads & "', '" & txtMnuSAMgr & "','admin_config_uploads.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",7"
arrData(9) = "'" & txtMnuSAMgrPM & "', '" & txtMnuSAMgr & "','admin_pm.asp','_parent','','','* Managers ADMIN','b_managers'," & pID & ",8"
arrData(10) = "'" & txtMnuSAMgrIPG & "', '" & txtMnuSAMgr & "','admin_ipgate.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",9"
arrData(11) = "'" & txtMnuSAMgrDBBack & "', '" & txtMnuSAMgr & "','admin_db.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",10"
arrData(12) = "'" & txtMnuSAMgrCtryFlgs & "', '" & txtMnuSAMgr & "','admin_countries.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",11"
arrData(13) = "'" & txtMnuSAMgrAV & "', '" & txtMnuSAMgr & "','admin_avatar_home.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",12"
arrData(14) = "'" & txtMnuSAMgrWelcome & "', '" & txtMnuSAMgr & "','admin_welcome.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",13"
arrData(15) = "'" & txtMnuSAMgrAnnce & "', '" & txtMnuSAMgr & "','admin_announce.asp','_parent','','','* " & txtMnuSAMgrAdmin & "','b_managers'," & pID & ",14"
populateB(arrData)
':: end MANAGERS button
end sub

sub admin_menu()
':: SUPERADMIN MENU BUTTONS ::::::::::::::::::::::::::::::
response.Write("<br />b_site_config<br />")
b_site_config()
response.Write("<br />b_mem_config<br />")
b_mem_config()
response.Write("<br />b_manager_config<br />")
b_manager_config()

response.Write("<br />b_avatar<br />")
b_avatar()
response.Write("<br />b_flags<br />")
b_flags()
response.Write("<br />b_banner_cfg<br />")
b_banner_cfg()
response.Write("<br />b_pm<br />")
b_pm()
response.Write("<br />b_layout_mgr<br />")
b_layout_mgr()
response.Write("<br />b_ipgate<br />")
b_ipgate()

response.Write("<br />sadmin<br />")
':: SUPERADMIN MENU ::::::::::::::::::::::::::::::
redim arrData(5)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'* " & txtMnuSACfgAdminMnu & "','sadmin','','','','','" & txtMnuSAPAdmin & "','sadmin',0,'b_site_cfg',1"
arrData(3) = "'* " & txtMnuSAMbrAdMnu & "','sadmin','','','','','" & txtMnuSAPAdmin & "','sadmin',0,'b_mem_cfg',2"
arrData(4) = "'* " & txtMnuSAMgrAdMnu & "','sadmin','','','','','" & txtMnuSAPAdmin & "','sadmin',0,'b_managers',3"
arrData(5) = "'* " & txtMnuSAModAdMnu & "','sadmin','','','','','" & txtMnuSAPAdmin & "','sadmin',0,'m_admin',4"
populateB(arrData)
':: end SUPERADMIN mneu ::::::::::::::::::::::::::::::::::::::::::::::::
end sub

'%%%%%%%%%% UPDATE FROM SkyPortal v RC2 to SkyPortal v RC3 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update_version(v)
	strSql = "UPDATE " & strTablePrefix & "APPS SET APP_VERSION = '" & v & "', APP_DATE = '" & DateToStr(now()) & "' WHERE APP_ID = 1"
	executeThis(strSql)
	
	strSql= "UPDATE " & strTablePrefix & "CONFIG SET C_PORTAL_VERSION = '" & v & "' WHERE CONFIG_ID = 1;"
	executeThis(strSql)
end sub

sub set_new_skin()
	':: reset the default site skin
	  ':: delete old default install skin
	  executeThis("delete from " & strTablePrefix & "COLORS where C_STRFOLDER = '" & itFolder & "'")
	  
	  ':: add to skins table
		strSql = "INSERT INTO " & strTablePrefix & "COLORS "
		strSql = strSql & "(C_STRFOLDER"
		strSql = strSql & ", C_STRDESCRIPTION"
		strSql = strSql & ", C_STRAUTHOR"
		strSql = strSql & ", C_TEMPLATE"
		strSql = strSql & ", C_STRTITLEIMAGE"
		strSql = strSql & ", C_INTSUBSKIN"
		strSql = strSql & ", C_SKINLEVEL"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & itFolder & "'"
		strSql = strSql & ", '" & itDesc & "'"
		strSql = strSql & ", '" & itAuthor & "'"
		strSql = strSql & ", '" & itName & "'"
		strSql = strSql & ", '" & itLogo & "'"
		strSql = strSql & ", " & itSubSkin
		strSql = strSql & ", '1,2,3'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		':: set default portal skin
		strSql = "UPDATE " & strTablePrefix & "CONFIG"
		strSql = strSql & " SET C_STRDEFTHEME = '" & itFolder & "'"
		strSql = strSql & ", C_INTSUBSKIN = " & itSubSkin
		strSql = strSql & ", C_STRTITLEIMAGE = '" & itLogo & "'"
		strSql = strSql & " WHERE CONFIG_ID = 1"
		executeThis(strSql)
		
		':: reset all member skins
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET THEME_ID = '0'"
		'strSql = strSql & " WHERE THEME_ID <> ''"
		executeThis(strSql)
end sub

%>