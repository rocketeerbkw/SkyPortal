<% 
'%%%%%%%% UPDATE FROM SkyPortal v RC7 to SkyPortal v RC8 %%%%%%%%%%%
sub update_rc7_v1()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC7 to SkyPortal vRC8</b><br /><br />")
  ':: rebuild avatar manager menu
  rc8_avatar_manager_menu()
  
  Response.Write("<br />Version Update<br />")
  update_version(longVer)
	
  mnu.DelMenuFiles("")
end sub

'%%%%%%%% UPDATE FROM SkyPortal v RC6 to SkyPortal v RC7 %%%%%%%%%%%
sub update_rc6_rc7()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC6 to SkyPortal vRC7</b><br /><br />")
  'Response.Write("<b>Add field</b><br />")
  sSql = "ALTER TABLE " & strTablePrefix & "M_SUBCATEGORIES ADD [ITEM_CNT] LONG DEFAULT 0"
  'alterTable(checkIt(strSql))
  'doSQL2 checkIt(sSql),1
  executeThis(sSql)
  sSql = "UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET ITEM_CNT=0"
  executeThis(sSql)
  
  sSql = "ALTER TABLE " & strTablePrefix & "M_SUBCATEGORIES ADD [S_SKYBLOCK] TEXT(100) NULL"
  'alterTable(checkIt(strSql))
  doSQL2 checkIt(sSql),1
  
  sSql = "ALTER TABLE " & strTablePrefix & "M_CATEGORIES ADD [C_SKYBLOCK] TEXT(100) NULL"
  'alterTable(checkIt(strSql))
  doSQL2 checkIt(sSql),1
  
 	':: add "Site Logs" link to "Managers" menu
	sSql = "select ID from menu where Parent = 'b_managers' and iName = 'b_managers'"
	set rsT = my_Conn.execute(sSql)
	pID = rsT(0)
	set rsT = nothing
	
	redim arrData(2)
	arrData(0) = "Menu"
	arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,iName,app_id,mnuAccess,mnuOrder,ParentID"
	arrData(2) = "'Site Logs', 'Managers','admin_logs.asp','','','_parent','* Managers ADMIN','b_managers',0,'1,2,3',1," & pID & ""
	populateB(arrData)
  
  'rc7_upload_config()
  rc7_new_indexes()
  
  Response.Write("<br />Version Update<br />")
  update_version(longVer)
	
  mnu.DelMenuFiles("")
end sub
  
'%%%%%%%% UPDATE FROM SkyPortal v RC5 to SkyPortal v RC6 %%%%%%%%%%%
sub update_rc5_rc6()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC5 to SkyPortal vRC6</b><br /><br />")
  Response.Write("<br />Add Config Fields<br />")
  
  'strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ADD [C_STREMAILPASSWORD] TEXT(255) NULL,[C_STREMAILUSERNAME] TEXT(255) NULL,[C_STREMAILPORT] INT NULL"
  'alterTable(checkIt(strSql))
	strSql = "" & strTablePrefix & "CONFIG,[C_STREMAILPASSWORD] TEXT(50) NULL,[C_STREMAILUSERNAME] TEXT(255) NULL,[C_STREMAILPORT] INT NULL"
	alterTable2(checkIt(strSql))
  
  strSql = "ALTER TABLE " & strTablePrefix & "UPLOAD_CONFIG DROP COLUMN [UP_ALLOWEDUSERS];" 
  'alterTable(checkIt(strSql))
  doSQL2 checkIt(strSql),1
  
  sSql = "UPDATE Menu SET mnuFunction = 'newPM' WHERE mnuFunction = 'pmImage'"
  executeThis(sSql)
  
  ':: Create the new integrated module tables
  createModuleTables()
  
  Response.Write("<br />Version Update<br />")
  update_version(longVer)
end sub

'%%%%%%%% UPDATE FROM SkyPortal v RC4 to SkyPortal v RC5 %%%%%%%%%%%
sub update_rc4_rc5()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC4 to SkyPortal vRC5</b><br />")
  Response.Write("<br />Version Update<br />")
  update_version(longVer)
end sub

'%%%%%%%%%% UPDATE FROM SkyPortal v RC3 to SkyPortal v RC4 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update_rc3_rc4()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC3 to SkyPortal vRC4</b><br />")
	'strSql = "ALTER TABLE " & strTablePrefix & "COLORS ALTER COLUMN C_SKINLEVEL TEXT(255)"
	'alterTable(checkIt(strSql))
	
    Response.Write("<br />Recreate Skin Table<br />")
	tblThemes()
	
    Response.Write("<br />Add App data<br />")
	strSQL = "UPDATE " & strTablePrefix & "APPS SET APP_tDATA1 = '1,2' WHERE APP_INAME = 'PM';"
	populateA(checkIt(strSql))
	
    Response.Write("<br />Add Comp Fields<br />")
	'strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ADD C_COMP_IMAGE TEXT(50), C_COMP_UPLOAD TEXT(50)"
	'doSQL2 checkIt(strSql),0
	
	strSql = "" & strTablePrefix & "CONFIG,[C_COMP_IMAGE] TEXT(50),[C_COMP_UPLOAD] TEXT(50)"
	alterTable2(checkIt(strSql))
	
    Response.Write("<br />Add Comp data<br />")
	strSql = "UPDATE " & strTablePrefix & "CONFIG SET C_COMP_IMAGE = 'NONE', C_COMP_UPLOAD = 'NONE' WHERE CONFIG_ID = 1;"
	populateA(checkIt(strSql))
	
  Response.Write("<br />Version Update<br />")
  update_version(longVer)
end sub

'%%%%%%%%%% UPDATE FROM SkyPortal v RC2 to SkyPortal v RC3 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update_rc2_rc3()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC2 to SkyPortal vRC3</b><br />")
  update_version(longVer)
  set_new_skin()
  
  sky_Pages()
  sky_menu()
end sub

'%%%%%%%%%% UPDATE FROM SkyPortal v RC1 to SkyPortal v RC2 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update_rc1_rc2()
  Response.Write("<hr /><b>UPDATE FROM SkyPortal vRC1 to SkyPortal vRC2</b><br />")

'PORTAL_APPS
	'add fields
	strSql = "" & strTablePrefix & "APPS,[APP_GROUPS_WRITE] MEMO NULL,[APP_GROUPS_FULL] MEMO NULL,[APP_VERSION] TEXT(10),[APP_DATE] TEXT(20),[APP_SUBSEC] INT"
	alterTable2(checkIt(strSql))
	err.clear
	
	strSql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_USERS = '1,2', APP_DATE = '" & DateToStr(now()) & "' WHERE APP_INAME = 'PM';"
	executeThis(strSql)
	
	strSql = "UPDATE " & strTablePrefix & "GROUPS SET G_NAME = 'Guests', G_INAME = 'Guests', G_DESC = 'Default group. All guest visitors.' WHERE G_INAME = 'Everyone';"
	executeThis(strSql)
	
	strSql = "" & strTablePrefix & "UPLOAD_CONFIG,[UP_ALLOWEDGROUPS] MEMO NULL"
	alterTable2(checkIt(strSql))
	
	strSql= "UPDATE " & strTablePrefix & "UPLOAD_CONFIG SET UP_ALLOWEDGROUPS = '1,2' WHERE ID not like 0;"
	executeThis(strSql)
	
	'add fields
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ADD C_PORTAL_VERSION TEXT(50)"
	doSQL2 checkIt(strSql),1
	
	strSql = "" & strTablePrefix & "CONFIG,[C_STRACTIVE] TEXT(50) NULL"
	alterTable2(checkIt(strSql))
	
	strSql= "UPDATE " & strTablePrefix & "CONFIG SET C_PORTAL_VERSION = '" & longVer & "' WHERE CONFIG_ID = 1;"
	executeThis(strSql)
	
	strSql= "UPDATE " & strTablePrefix & "FP SET FP_GROUPS = '1,2,3' WHERE id NOT LIKE 0;"
	executeThis(strSql)
	
	set_new_skin()
		
	err.clear
	
end sub

'%%%%%%%%%% UPDATE FROM MWP v2.1 to SkyPortal v RC1 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
sub update_211xRC1()
  Response.Write("<hr /><b>UPDATE FROM MWP.info v2.1 to SkyPortal vRC1</b>")
    tblAPPS()
	tblThemes()
	tblUploads()
    cr_new_SP_RC1()
	
'PORTAL_CP_CONFIG
	'add fields
	strSql = "" & strTablePrefix & "CP_CONFIG,[PM_OUTBOX] INT"
	alterTable2(checkIt(strSql))
	err.clear
	
	strSql = "UPDATE " & strTablePrefix & "CP_CONFIG SET PM_OUTBOX = 1 WHERE MEMBER_ID <> 0"
	populateA(checkIt(strSql))

'PORTAL_MEMBERS
	'add fields
	strSql = "" & strMemberTablePrefix & "MEMBERS, [M_LCID] LONG, [M_TIME_OFFSET] LONG DEFAULT 0, [M_TIME_TYPE] TEXT(2), [M_PMSTATUS] INT DEFAULT 1, [M_PMBLACKLIST] MEMO, [M_DONATE] LONG DEFAULT 0, [M_LANG] TEXT(2) NULL"
	alterTable2(checkIt(strSql))
	err.clear
	
	strSQL = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_LCID=" & intPortalLCID & ", M_TIME_OFFSET=" & timeoffset & ",M_TIME_TYPE='12',THEME_ID ='0',M_DONATE = 0, M_PMSTATUS = 1, M_LANG = 'en' WHERE MEMBER_ID <> 0"
	populateA(checkIt(strSql))
	
	strSql = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ALTER COLUMN M_SIG MEMO"
	alterTable(checkIt(strSql))
	err.clear

'PORTAL_MEMBERS_PENDING
	strSql = "" & strMemberTablePrefix & "MEMBERS_PENDING, M_LCID LONG, M_TIME_OFFSET LONG DEFAULT 0, M_TIME_TYPE TEXT(50), [M_LANG] TEXT(2) NULL"
	alterTable2(checkIt(strSql))
	
	strSQL = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING SET M_LCID=" & intPortalLCID & ", M_TIME_OFFSET=0,M_TIME_TYPE='12', M_LANG = 'en' WHERE M_NAME <> ''"
	populateA(checkIt(strSql))

'PORTAL_PM
	'add fields
	strSql = "" & strTablePrefix & "PM,[M_SAVED] INT DEFAULT 0"
	alterTable2(checkIt(strSql))
	err.clear
	
	
	strSQL = "UPDATE " & strTablePrefix & "PM SET M_SAVED=0 WHERE M_ID <> 0"
	populateA(checkIt(strSql))
	
	strSql = "ALTER TABLE " & strTablePrefix & "PM ALTER COLUMN M_SUBJECT TEXT(100)"
	alterTable(checkIt(strSql))
	err.clear
	
'PORTAL_ONLINE - 
	droptable("" & strMemberTablePrefix & "ONLINE")
	sSQL = "CREATE TABLE [" & strMemberTablePrefix & "ONLINE]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CheckedIn] TEXT(100), [DateCreated] TEXT(100), [LastChecked] TEXT(100), [LastDateChecked] TEXT(100), [M_BROWSE] MEMO, [UserID] TEXT(100), [UserIP] TEXT(40), [UserAgent] TEXT(40));"

	createTable(checkIt(sSQL))
	
'PORTAL_CONFIG - 
	'change text length
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ALTER COLUMN C_STRCOPYRIGHT TEXT(200)"
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ALTER COLUMN C_STRVAR1 TEXT(50)"
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ALTER COLUMN C_STRVAR2 TEXT(50)"
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ALTER COLUMN C_STRVAR3 TEXT(50)"
	alterTable(checkIt(strSql))
	err.clear
	
	'::add fields
	strSql = "" & strTablePrefix & "CONFIG,[C_INTSUBSKIN] INT,[C_ONEADAYDATE] TEXT(20),[C_STRVAR8] TEXT(50),[C_STRVAR9] TEXT(50),[C_VERSION] TEXT(20)"
	alterTable2(checkIt(strSql))
	err.clear
	
	strSQL = "UPDATE " & strTablePrefix & "CONFIG SET C_STRDEFTHEME='" & installTheme & "'"
		strSql = strSql & ", C_INTSUBSKIN=1"
		strSql = strSql & ", C_ONEADAYDATE='" & DateToStr(date()) & "'"
		strSql = strSql & ", C_VERSION='" & longVer & "'"
		strSql = strSql & " WHERE CONFIG_ID = 1"
	populateA(checkIt(strSql))

	':: drop fields no longer used
   response.Write("<hr /><h5>Drop fields no longer used</h5>")
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_MODULES]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_PAGEWIDTH]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRACTIVELINKCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRALTFORUMCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRALTHEADCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRCATEGORYCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRCATEGORYFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRDEFAULTFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRDEFAULTFONTFACE]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRDEFAULTFONTSIZE]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFOOTERFONTSIZE]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFORUMCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFORUMFIRSTCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFORUMFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFORUMLINKCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRFORUMURL]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRHEADCELLCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRHEADERFONTSIZE]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRHEADFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRHOVERFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRHOVERTEXTDECORATION]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRLINKCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRLINKTEXTDECORATION]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRNEWFONTCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRPAGEBGCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRPAGEBGIMAGE]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRPOPUPBORDERCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRPOPUPTABLECOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRSETCOOKIETOFORUM]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRTABLEBORDERCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRTOPICNOWRAPLEFT]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRTOPICNOWRAPRIGHT]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRTOPICWIDTHLEFT]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRTOPICWIDTHRIGHT]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRVISITEDLINKCOLOR]" 
	alterTable(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG DROP COLUMN [C_STRVISITEDTEXTDECORATION]" 
	alterTable(checkIt(strSql))
	err.clear

'-------------------- populate table with default values --------------------------
   response.Write("<hr /><h5>Banners default data</h5>")
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
		strSql = strSql & ", '" & portalUrl & "files/banners/SkyPortal.gif'"
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
		strSql = strSql & ", '" & portalUrl & "files/banners/webdogg.jpg'"
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
		strSql = strSql & ", '" & portalUrl & "files/banners/liveair.gif'"
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
		strSql = strSql & ", '" & portalUrl & "files/banners/aff_SkyPortal.gif'"
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
		strSql = strSql & ", '" & portalUrl & "files/banners/aff_webdogg.gif'"
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
		strSql = strSql & ", '" & portalUrl & "files/banners/aff_liveair.gif'"
		strSql = strSql & ", 2"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
	
':: End of core.

  response.Write("<hr />")
':: Start Module upgrades.
	spArticles()
	spClassifieds()
	spDownloads()
	spForums()
	spLinks()
	spPictures()
  response.Write("<hr />")
    
end sub

sub rc8_avatar_manager_menu()
  ':: start button template
  mnuName = "* " & txtMnuAvAdmin	
  mnuINAME = "b_avatar_cfg"
  mnuBName = txtMnuAvSetup

  'redim arrData(2)
  'arrData(0) = "Menu"
  'arrData(1) = "Name,Parent,Link,mnuImage,onclick,Target,mnuTitle,INAME,mnuAccess,mnuOrder"
  'arrData(2) = "'" & mnuBName & "', '" & mnuINAME & "','','','','','" & mnuName & "','" & mnuINAME & "','',1"
  'populateB(arrData)
  
  sSql = "DELETE FROM menu WHERE Parent = '" & mnuBName & "' and INAME = '" & mnuINAME & "'"
  executeThis(sSql)

  sSql = "select ID from menu where Parent = '" & mnuINAME & "' and INAME = '" & mnuINAME & "'"
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

sub rc7_clean_members()
  Response.Write("<br /><b>Clean members table</b><br />")
  sSql = "DELETE FROM PORTAL_MEMBERS WHERE M_EMAIL = ''"
  executeThis(sSql)
  sSql = "DELETE FROM PORTAL_MEMBERS WHERE M_EMAIL = ' '"
  executeThis(sSql)
end sub
  
sub rc7_upload_config()
  Response.Write("<br /><b>Modify Upload Config</b><br />")
  
  sSql = "DELETE FROM " & strTablePrefix & "UPLOAD_CONFIG"
  sSql = sSql & " WHERE UP_LOCATION = 'avatar'"
  executeThis(sSql)
  
  sSql = "DELETE FROM " & strTablePrefix & "UPLOAD_CONFIG"
  sSql = sSql & " WHERE UP_LOCATION = 'photo'"
  executeThis(sSql)
  
	strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
	strSql = strSql & "(UP_SIZELIMIT"
	strSql = strSql & ", UP_ALLOWEDEXT"
	strSql = strSql & ", UP_LOGUSERS"
	strSql = strSql & ", UP_ALLOWEDGROUPS"
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
	strSql = strSql & ", '1,2'"
	strSql = strSql & ", 'Member Photo'"
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
	'response.Write(strSql)
	populateA(strSql)

	strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
	strSql = strSql & "(UP_SIZELIMIT"
	strSql = strSql & ", UP_ALLOWEDEXT"
	strSql = strSql & ", UP_LOGUSERS"
	strSql = strSql & ", UP_ALLOWEDGROUPS"
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
	strSql = strSql & ", '1,2'"
	strSql = strSql & ", 'Member Avatar'"
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
	'response.Write(strSql)
	populateA(strSql)

	strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
	strSql = strSql & "(UP_SIZELIMIT"
	strSql = strSql & ", UP_ALLOWEDEXT"
	strSql = strSql & ", UP_LOGUSERS"
	strSql = strSql & ", UP_ALLOWEDGROUPS"
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
	strSql = strSql & ", '1'"
	strSql = strSql & ", 'Portal Avatars'"
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
	'response.Write(strSql)
	populateA(strSql)
end sub

sub rc7_new_indexes()
  Response.Write("<br /><b>Create Module Table indexes</b><br />")
  redim indexes(5)
  indexes(0) = "CREATE INDEX [MC_PARENT_ID] ON [" & strTablePrefix & "M_CATEGORIES]([PARENT_ID]);"
  indexes(1) = "CREATE INDEX [MC_APP_ID] ON [" & strTablePrefix & "M_CATEGORIES]([APP_ID]);"
  indexes(2) = "CREATE INDEX [MS_CAT_ID] ON [" & strTablePrefix & "M_SUBCATEGORIES]([CAT_ID]);"
  indexes(3) = "CREATE INDEX [MS_APP_ID] ON [" & strTablePrefix & "M_SUBCATEGORIES]([APP_ID]);"
  indexes(4) = "CREATE INDEX [MR_ITEM_ID] ON [" & strTablePrefix & "M_RATING]([ITEM_ID]);"
  indexes(5) = "CREATE INDEX [MR_APP_ID] ON [" & strTablePrefix & "M_RATING]([APP_ID]);"
  createIndx(indexes)
end sub

sub createModuleTables()
  ':: create parent table
  response.Write("<hr><b>Create " & strTablePrefix & "M_PARENT table</b><br /><br />")
  sSQL = "CREATE TABLE [" & strTablePrefix & "M_PARENT]("
  sSQL = sSql & "[PARENT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL"
  sSQL = sSql & ", [PARENT_NAME] TEXT(100)"
  sSQL = sSql & ", [PARENT_SDESC] MEMO NULL"
  sSQL = sSql & ", [PARENT_LDESC] MEMO NULL"
  sSQL = sSql & ", [PARENT_IMAGE] TEXT(255) NULL"
  sSQL = sSQL & ", [PG_READ] MEMO NULL"
  sSQL = sSQL & ", [PG_WRITE] MEMO NULL"
  sSQL = sSQL & ", [PG_FULL] MEMO NULL"
  sSQL = sSQL & ", [PG_INHERIT] INT DEFAULT 1"
  sSQL = sSQL & ", [PG_PROPAGATE] INT DEFAULT 1"
  sSQL = sSql & ", [APP_ID] LONG"
  sSQL = sSql & ", [C_ORDER] INT DEFAULT 1"
  sSQL = sSql & ");"
  createTable(checkIt(sSQL))
  
  sSql = "INSERT INTO " & strTablePrefix & "M_PARENT ("
  sSQL = sSql & "PARENT_NAME,PARENT_SDESC,PARENT_LDESC"
  sSQL = sSql & ",PG_READ,PG_WRITE,PG_FULL,PG_INHERIT"
  sSQL = sSql & ",PG_PROPAGATE,APP_ID"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'Portal Parent Group'"
  sSQL = sSql & ",'Default initial parent group'"
  sSQL = sSql & ",'Default initial parent group."
  sSQL = sSql & " You cannot delete this group'"
  sSQL = sSql & ",'1,2,3','1,2','1',1,1,0"
  sSQL = sSql & ")"
  executeThis(sSQL)
  
  ':: create category table
  response.Write("<hr><b>Create " & strTablePrefix & "M_CATEGORIES table</b><br /><br />")
  sSQL = "CREATE TABLE [" & strTablePrefix & "M_CATEGORIES]("
  sSQL = sSql & "[CAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL"
  sSQL = sSql & ", [PARENT_ID] LONG NULL DEFAULT 1"
  sSQL = sSql & ", [CAT_NAME] TEXT(100)"
  sSQL = sSql & ", [CAT_SDESC] MEMO NULL"
  sSQL = sSql & ", [CAT_LDESC] MEMO NULL"
  sSQL = sSql & ", [CAT_IMAGE] TEXT(255) NULL"
  sSQL = sSQL & ", [CG_READ] MEMO NULL"
  sSQL = sSQL & ", [CG_WRITE] MEMO NULL"
  sSQL = sSQL & ", [CG_FULL] MEMO NULL"
  sSQL = sSQL & ", [CG_INHERIT] INT DEFAULT 1"
  sSQL = sSQL & ", [CG_PROPAGATE] INT DEFAULT 1"
  sSQL = sSql & ", [APP_ID] LONG"
  sSQL = sSql & ", [C_ORDER] INT DEFAULT 1"
  sSQL = sSql & ");"
  createTable(checkIt(sSQL))
  
  ':: create subcategory table
  response.Write("<hr><b>Create " & strTablePrefix & "M_SUBCATEGORIES table</b><br /><br />")
  sSQL = "CREATE TABLE [PORTAL_M_SUBCATEGORIES]("
  sSQL = sSql & "[SUBCAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL"
  sSQL = sSql & ", [CAT_ID] LONG"
  sSQL = sSql & ", [SUBCAT_NAME] TEXT(100)"
  sSQL = sSql & ", [SUBCAT_SDESC] MEMO NULL"
  sSQL = sSql & ", [SUBCAT_LDESC] MEMO NULL"
  sSQL = sSql & ", [SUBCAT_IMAGE] TEXT(255) NULL"
  sSQL = sSQL & ", [SG_READ] MEMO NULL"
  sSQL = sSQL & ", [SG_WRITE] MEMO NULL"
  sSQL = sSQL & ", [SG_FULL] MEMO NULL"
  sSQL = sSQL & ", [SG_INHERIT] INT DEFAULT 1"
  sSQL = sSql & ", [APP_ID] LONG"
  sSQL = sSql & ", [C_ORDER] INT DEFAULT 1"
  sSQL = sSql & ");"
  createTable(checkIt(sSQL))
  
  ':: create rating table
  response.Write("<hr><b>Create " & strTablePrefix & "M_RATING table</b><br /><br />")
  sSQL = "CREATE TABLE [" & strTablePrefix & "M_RATING]("
  sSQL = sSql & "[RATING_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL"
  sSQL = sSql & ", [ITEM_ID] LONG"
  sSQL = sSql & ", [RATE_BY] LONG"
  sSQL = sSql & ", [RATE_DATE] TEXT(50)"
  sSQL = sSql & ", [RATING] INT"
  sSQL = sSql & ", [COMMENTS] MEMO"
  sSQL = sSql & ", [APP_ID] LONG"
  sSQL = sSql & ");"
  createTable(checkIt(sSQL))
end sub

sub migrate_mwpx()
    Response.Write("<br />Alter Portal_CONFIG<br />")
	strSql = "ALTER TABLE " & strTablePrefix & "CONFIG ADD [C_MODULES] TEXT(50), [C_SECIMAGE] INT, [C_STRLOCKDOWN] INT"
	doSQL2 checkIt(strSql),0
	
    Response.Write("<br />Add Portal_CONFIG default data<br />")
	strSql = "UPDATE " & strTablePrefix & "CONFIG SET"
	strSql = strSql & " C_MODULES = '0'"
	strSql = strSql & ", C_SECIMAGE = 0"
	strSql = strSql & ", C_STRLOCKDOWN = 0"
	strSql = strSql & ", C_STRDEFTHEME = '" & installTheme & "'"
	
	strSql = strSql & " WHERE CONFIG_ID = 1;"
	populateA(checkIt(strSql))
	
    Response.Write("<br />Alter Portal_CP_CONFIG<br />")
	strSql = "ALTER TABLE " & strTablePrefix & "CP_CONFIG ADD THEME_ID TEXT(50)"
	doSQL2 checkIt(strSql),0
	
    Response.Write("<br />Add Portal_CP_CONFIG default data<br />")
	strSql = "UPDATE " & strTablePrefix & "CP_CONFIG SET THEME_ID = '0' WHERE ID <> 0;"
	populateA(checkIt(strSql))
	
    Response.Write("<br />Alter Portal_members_pending<br />")
	strSql = "ALTER TABLE " & strTablePrefix & "MEMBERS_PENDING ADD THEME_ID TEXT(50)"
	doSQL2 checkIt(strSql),0
	
    Response.Write("<br />Add Portal_members_pending default data<br />")
	strSql = "UPDATE " & strTablePrefix & "MEMBERS_PENDING SET THEME_ID = '0' WHERE MEMBER_ID <> 0;"
	populateA(checkIt(strSql))
end sub

sub spArticles()
  response.Write("<hr /><h4>ARTICLES</h4>")
 if buArticle = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'article','article',1,0,'1,3',1,1,'config_articles','article.asp'"
  populateB(arrData)
  
  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'article'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing

	'::::::::::::::::::::: CREATE ARTICLE FRONT PAGE ITEMS :::::::::::::::::::::::::
  response.Write("<br /><h5>Update " & strTablePrefix & "FP table with ARTICLE info</h5>")
	'articles front page
	redim arrData(9)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Articles - Most Viewed','a_popular_sm','article_sm:top',1,4,'Most popular articles by hit count.','3'," & app_id & ""
	arrData(3) = "'Articles - Most Viewed','a_popular_lg','article_lg:top',1,2,'Most popular articles by hit count.','3'," & app_id & ""
	arrData(4) = "'Articles - Newest','a_newest_sm','article_sm:new',1,4,'Newest articles.','3'," & app_id & ""
	arrData(5) = "'Articles - Newest','a_newest_lg','article_lg:new',1,2,'Newest articles.','3'," & app_id & ""
	arrData(6) = "'Articles - Featured','a_admin_sm','article_sm:featured',1,4,'Front page articles specified by admin.','3'," & app_id & ""
	arrData(7) = "'Articles - Featured','a_admin_lg','article_lg:featured',1,2,'Front page articles specified by admin.','3'," & app_id & ""
	arrData(8) = "'Articles - Random','a_rand_sm','article_sm:rand',1,4,'Front page articles.','3'," & app_id & ""
	arrData(9) = "'Articles - Random','a_rand_lg','article_lg:rand',1,2,'Front page articles.','3'," & app_id & ""
	populateB(arrData)
	
    response.Write("<br /><h5>Add fields to existing Articles tables</h5>")
	'add fields
	strSql = "ARTICLE, FEATURED BIT DEFAULT 0"
	alterTable2(checkIt(strSql))
	err.clear
	strSql = "UPDATE ARTICLE SET FEATURED=0 WHERE ARTICLE_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "ARTICLE_CATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE ARTICLE_CATEGORIES SET C_ORDER=1 WHERE CAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "ARTICLE_SUBCATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE ARTICLE_SUBCATEGORIES SET C_ORDER=1 WHERE SUBCAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
    response.Write("<br /><h5>Articles Updated</h5>")
 else
 	droptable("ARTICLE")
 	droptable("ARTICLE_CATEGORIES")
 	droptable("ARTICLE_SUBCATEGORIES")
 	droptable("ARTICLE_RATING")
    response.Write("<br /><h5>Articles Deleted</h5>")
 end if	
  ':: end article upgrade
end sub

sub spClassifieds()
' buClassified
  response.Write("<hr /><h4>CLASSIFIEDS</h4>")
 if buClassified = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'classifieds','classifieds',1,0,'1,3',1,1,'config_classifieds','classified.asp'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'classifieds'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
	
	'add classifieds front page items
  response.Write("<br /><h5>Add data to " & strTablePrefix & "FP table</h5>")
	redim arrData(7)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Classifieds - Newest','c_newest_sm','class_sm:new',1,4,'Newest classifieds.','3'," & app_id & ""
	arrData(3) = "'Classifieds - Newest','c_newest_lg','class_lg:new',1,2,'Newest classifieds.','3'," & app_id & ""
	arrData(4) = "'Classifieds - Featured','c_admin_sm','class_sm:featured',1,4,'Front page classifieds specified by admin.','3'," & app_id & ""
	arrData(5) = "'Classifieds - Featured','c_admin_lg','class_lg:featured',1,2,'Front page classifieds specified by admin.','3'," & app_id & ""
	arrData(6) = "'Classifieds - Random','c_rand_sm','class_sm:rand',1,4,'Front page random classifieds.','3'," & app_id & ""
	arrData(7) = "'Classifieds - Random','c_rand_lg','class_lg:rand',1,2,'Front page random classifieds.','3'," & app_id & ""
	populateB(arrData)

  response.Write("<br /><h5>Add data to " & strTablePrefix & "UPLOAD_CONFIG table</h5>")
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
		strSql = strSql & "100"
		strSql = strSql & ", 'gif,jpg'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 3"
		strSql = strSql & ", 'classified'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 300"
		strSql = strSql & ", 300"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/classified_images/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
	
  response.Write("<br /><h5>Add fields to existing Classifieds tables</h5>")
	'add fields
	strSql = "CLASSIFIED, FEATURED BIT DEFAULT 0"
	alterTable2(checkIt(strSql))
	err.clear
	strSql = "UPDATE CLASSIFIED SET FEATURED=0 WHERE CLASSIFIED_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "CLASSIFIED_CATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE CLASSIFIED_CATEGORIES SET C_ORDER=1 WHERE CAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "CLASSIFIED_SUBCATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE CLASSIFIED_SUBCATEGORIES SET C_ORDER=1 WHERE SUBCAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
    response.Write("<br /><h5>CLASSIFIEDS Updated</h5>")
 else
 	droptable("CLASSIFIED")
 	droptable("CLASSIFIED_CATEGORIES")
 	droptable("CLASSIFIED_SUBCATEGORIES")
    response.Write("<br /><h5>CLASSIFIEDS Deleted</h5>")
 end if
	
  ':: end classified upgrade
end sub

sub spDownloads()
'  buDL
  response.Write("<hr /><h4>DOWNLOADS</h4>")
 if buDL = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'downloads','downloads',1,0,'1,3',1,1,'config_downloads','dl.asp'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'downloads'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
	
	'add downloads to front page items
  response.Write("<br /><h5>Add data to " & strTablePrefix & "FP table<'h5>")
	redim arrData(9)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Downloads - Popular','dl_popular_sm','dl_small:top',1,4,'Most popular downloads.','3'," & app_id & ""
	arrData(3) = "'Downloads - Popular','dl_popular_lg','dl_large:top',1,2,'Most popular downloads.','3'," & app_id & ""
	arrData(4) = "'Downloads - Newest','dl_newest_sm','dl_small:new',1,4,'Newest downloads.','3'," & app_id & ""
	arrData(5) = "'Downloads - Newest','dl_newest_lg','dl_large:new',1,2,'Newest downloads.','3'," & app_id & ""
	arrData(6) = "'Downloads - Random','dl_random_sm','dl_small:random',1,4,'Random downloads.','3'," & app_id & ""
	arrData(7) = "'Downloads - Random','dl_random_lg','dl_large:random',1,2,'Random downloads.','3'," & app_id & ""
	arrData(8) = "'Downloads - Featured','dl_featured_sm','dl_small:featured',1,4,'Featured downloads.','3'," & app_id & ""
	arrData(9) = "'Downloads - Featured','dl_featured_lg','dl_large:featured',1,2,'Featured downloads.','3'," & app_id & ""
	populateB(arrData)

  response.Write("<br /><h5>Add data to " & strTablePrefix & "UPLOAD_CONFIG table</h5>")
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
		strSql = strSql & ", 'zip,rar'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 3"
		strSql = strSql & ", 'download'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/downloads/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
	
  response.Write("<br /><h5>Add fields to existing Downloads tables</h5>")
	'add fields
	strSql = "DL, FEATURED BIT DEFAULT 0"
	alterTable2(checkIt(strSql))
	err.clear
	strSql = "UPDATE DL SET FEATURED=0 WHERE DL_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "DL_CATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE DL_CATEGORIES SET C_ORDER=1 WHERE CAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "DL_SUBCATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE DL_SUBCATEGORIES SET C_ORDER=1 WHERE SUBCAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
    response.Write("<br /><h5>DOWNLOADS Updated</h5>")
 else
 	droptable("DL")
 	droptable("DL_CATEGORIES")
 	droptable("DL_SUBCATEGORIES")
 	droptable("DL_RATING")
    response.Write("<br /><h5>DOWNLOADS Deleted</h5>")
 end if
  ':: end Downloads upgrade
end sub

sub spForums()
'  buForums
  response.Write("<hr /><h4>FORUMS</h4>")
 if buForums = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'forums','forums',1,0,'1,3,4',1,1,'config_forums','link.asp?topicID='"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'forums'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing

':::::::::::: add forum items to front page :::::::::::::::::::::::::::
	response.Write("<br /><h5>Add forum items to front page</h5>")
	redim arrData(4)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Forum recent topics','forum_topics','f_topics_sm',1,4,'Recent topics from the forum.','3'," & app_id & ""
	arrData(3) = "'Featured Polls','forum_polls','f_polls_fp',1,4,'Featured polls from the forums.','3'," & app_id & ""
	arrData(4) = "'Site News','forum_news','f_news_fp',1,2,'Website News.','3'," & app_id & ""
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
	
    response.Write("<br /><h5>FORUMS Updated</h5>")
 else
    response.Write("<br /><h5>FORUMS Deleted</h5>")
 end if
  ':: end Forums upgrade
end sub

sub spPictures()
'  buPics
  response.Write("<hr /><h4>PICTURES</h4>")
 if buPics = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'pictures','pictures',1,0,'1,3',1,1,'config_pictures','pic.asp'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'pictures'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
	
  response.Write("<br /><h5>Add data to " & strTablePrefix & "FP table</h5>")
	redim arrData(9)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
'photos
	arrData(2) = "'Pictures - Most viewed','p_popular_sm','photos_sm:top',1,4,'Most popular photos.','3'," & app_id & ""
	arrData(3) = "'Pictures - Most viewed','p_popular_lg','photos_lg:top',1,2,'Most popular photos.','3'," & app_id & ""
	arrData(4) = "'Pictures - Newest','p_newest_sm','photos_sm:new',1,4,'Newest photos.','3'," & app_id & ""
	arrData(5) = "'Pictures - Newest','p_newest_lg','photos_lg:new',1,2,'Newest photos.','3'," & app_id & ""
	arrData(6) = "'Pictures - Featured','p_featured_sm','photos_sm:featured',1,4,'Featured photos specified by admin.','3'," & app_id & ""
	arrData(7) = "'Pictures - Featured','p_featured_lg','photos_lg:featured',1,2,'Featured photos specified by admin.','3'," & app_id & ""
	arrData(8) = "'Pictures - Random','p_rand_sm','photos_sm:rand',1,4,'Random photos.','3'," & app_id & ""
	arrData(9) = "'Pictures - Random','p_rand_lg','photos_lg:rand',1,2,'Random photos.','3'," & app_id & ""
	populateB(arrData)

'::::::::: upload config :::::::::::::::::::::::::::::
  response.Write("<br /><h5>Add data to " & strTablePrefix & "UPLOAD_CONFIG table</h5>")
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
		strSql = strSql & ", 'gif,jpg,png'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 3"
		strSql = strSql & ", 'pictures'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 120"
		strSql = strSql & ", 120"
		strSql = strSql & ", 500"
		strSql = strSql & ", 500"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'files/pictures/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
	
  response.Write("<br /><h5>Add fields to existing Pictures tables</h5>")
	'add fields
	strSql = "PIC, FEATURED BIT DEFAULT 0"
	alterTable2(checkIt(strSql))
	err.clear
	strSql = "ALTER TABLE PIC ALTER COLUMN OWNER TEXT(255)"
	alterTable(checkIt(strSql))
	err.clear
	strSql = "UPDATE PIC SET FEATURED=0 WHERE PIC_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "PIC_CATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE PIC_CATEGORIES SET C_ORDER=1 WHERE CAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "PIC_SUBCATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE PIC_SUBCATEGORIES SET C_ORDER=1 WHERE SUBCAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
  
    response.Write("<br /><h5>PICTURES Updated</h5>")
 else
 	droptable("PIC")
 	droptable("PIC_CATEGORIES")
 	droptable("PIC_SUBCATEGORIES")
 	droptable("PIC_RATING")
    response.Write("<br /><h5>PICTURES Deleted</h5>")
 end if
  ':: end pictures upgrade
end sub

sub spLinks()
'  buLinks
  response.Write("<hr /><h4>LINKS</h4>")
 if buLinks = 1 then
  response.Write("<h5>Update " & strTablePrefix & "APPS</h5>")
  'create the app
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW]"
  arrData(2) = "'links','links',1,0,'1,3',1,1,'config_links','links_pop.asp'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'links'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
	
'-------------------- populate FRONT PAGE table with default values ------------------
  response.Write("<br /><h5>Add data to " & strTablePrefix & "FP table</h5>")
	'add links to front page otems
	redim arrData(9)
	arrData(0) = "[" & strTablePrefix & "FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Links - Most viewed','l_popular_sm','links_sm:top',1,4,'Most popular links.','3'," & app_id & ""
	arrData(3) = "'Links - Most viewed','l_popular_lg','links_lg:top',1,2,'Most popular links.','3'," & app_id & ""
	arrData(4) = "'Links - Newest','l_newest_sm','links_sm:new',1,4,'Newest links.','3'," & app_id & ""
	arrData(5) = "'Links - Newest','l_newest_lg','links_lg:new',1,2,'Newest links.','3'," & app_id & ""
	arrData(6) = "'Links - Featured','l_featured_sm','links_sm:featured',1,4,'Featured links specified by admin.','3'," & app_id & ""
	arrData(7) = "'Links - Featured','l_featured_lg','links_lg:featured',1,2,'Featured links specified by admin.','3'," & app_id & ""
	arrData(8) = "'Links - Random','l_rand_sm','links_sm:rand',1,4,'Front page random links.','3'," & app_id & ""
	arrData(9) = "'Links - Random','l_rand_lg','links_lg:rand',1,2,'Front page random links.','3'," & app_id & ""
	populateB(arrData)

  response.Write("<br /><h5>Add data to " & strTablePrefix & "UPLOAD_CONFIG table</h5>")
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
		strSql = strSql & "100"
		strSql = strSql & ", 'gif,jpg'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 3"
		strSql = strSql & ", 'link'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 300"
		strSql = strSql & ", 300"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'files/link_images/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
  response.Write("<br /><h5>Add fields to existing Links tables</h5>")
	'add fields
	strSql = "LINKS, FEATURED BIT DEFAULT 0, POSTER TEXT(100)"
	alterTable2(checkIt(strSql))
	err.clear
	strSql = "UPDATE LINKS SET FEATURED=0 WHERE LINK_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "LINKS_CATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE LINKS_CATEGORIES SET C_ORDER=1 WHERE CAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
	strSql = "LINKS_SUBCATEGORIES, C_ORDER INT DEFAULT 1,GROUPS MEMO NULL"
	alterTable2(checkIt(strSql))
	strSql = "UPDATE LINKS_SUBCATEGORIES SET C_ORDER=1 WHERE SUBCAT_ID NOT LIKE 0"
	executeThis(checkIt(strSql))
	err.clear
	
    response.Write("<br /><h5>LINKS Updated</h5>")
 else
 	droptable("LINKS")
 	droptable("LINKS_CATEGORIES")
 	droptable("LINKS_SUBCATEGORIES")
 	droptable("LINKS_RATING")
    response.Write("<br /><h5>LINKS Deleted</h5>")
 end if
	
  ':: end WebLinks upgrade
end sub
%>