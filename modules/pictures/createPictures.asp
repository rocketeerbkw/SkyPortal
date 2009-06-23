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

':: leave this as is. Edit this value in fp_pic.asp
strPicTablePrefix = ""
app_version = "0.11"
do_app = true
incPicFp = false
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
  if incPicFp then
    createPictures()
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the fp_pic.asp ""include"" file<br>")
    Response.Write("to your fp_custom.asp file in order<br>")
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

sub pic_Upgrades()
  pic_Upgrade09()
  pictures_Upgrade_10_11()
end sub

sub createPictures()
  spThemeBlock1_open(intSkin)
  response.Write("<hr><b>PICTURES MODULE</b><br><br>")
  'check if app is existing
  sSql = "SELECT APP_ID, APP_NAME, APP_VERSION FROM PORTAL_APPS WHERE APP_INAME = 'pictures'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall or bReinstall then
      uninstall_Pictures()
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing
  
  if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "0.11"
	   ':: upcoming release
     case "0.10"
	   ':: current version
	   pictures_Upgrade_10_11()
     case "0.9"
	   ':: no changes, just update to version 0.10
	   updateVersion app_version,"pictures"
	 case else
	   pic_Upgrades()
   end select
  elseif not bUninstall then
    addPicApp()
    addPicFp()
    addPicUploads()
  
    crPicCatTbl()
    crPicSubCatTbl()
    crPicMainTbl()
    crPicRatingTbl()

    ':: do the upgrades
    pic_Upgrades()
  end if
 if not bUninstall then
  response.Write("<hr><h3>" & strPicTablePrefix & "Pictures Module Installed</h3><br><br>")
 else
  response.Write("<hr><h3>" & strPicTablePrefix & "Pictures Module Uninstalled</h3><br><br>")
 end if
  response.Write("<b>Be sure to delete this file (createPictures.asp) from your server!</b><br><br>")
  response.Write("<a href=""pic.asp""><b>Continue</b></a><br><br><br><br>")
  spThemeBlock1_close(intSkin)
  Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
end sub


sub addPicApp()
  'create the app
  response.Write("<hr><b>Update PORTAL_APPS table</b><br><br>")
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW],[APP_VERSION]"
  arrData(2) = "'" & strPicTablePrefix & "pictures','" & strPicTablePrefix & "pictures',1,0,'1,2,3',1,1,'config_pictures','pic.asp','" & app_version & "'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM PORTAL_APPS WHERE APP_INAME = '" & strPicTablePrefix & "pictures'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub addPicFp()
  response.Write("<hr><b>Add data to PORTAL_FP table</b><br><br>")
	redim arrData(9)
	arrData(0) = "[PORTAL_FP]"
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
end sub

sub addPicUploads()
  '::::::::: upload config :::::::::::::::::::::::::::::
  response.Write("<hr><b>Add data to PORTAL_UPLOAD_CONFIG table</b><br><br>")
		strSql = "INSERT INTO PORTAL_UPLOAD_CONFIG "
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
		strSql = strSql & ", 'gif,jpg,png'"
		strSql = strSql & ", 1"
		strSql = strSql & ", '1,2'"
		strSql = strSql & ", '" & strPicTablePrefix & "pictures'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.txt'"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 120"
		strSql = strSql & ", 120"
		strSql = strSql & ", 500"
		strSql = strSql & ", 500"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'files/" & strPicTablePrefix & "pictures/'"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub crPicCatTbl()
  response.Write("<hr><b>Create PIC_CATEGORIES table</b><br><br>")
  '::::::::::::::::::::::: CREATE PIC_CATEGORIES TABLE :::::::::::::::::::::::::::::
  sSQL = "CREATE TABLE [" & strPicTablePrefix & "PIC_CATEGORIES]([CAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CAT_IMAGE] TEXT(100), [CAT_NAME] TEXT(100),[C_ORDER] INT DEFAULT 1,[GROUPS] MEMO);"

  createTable(checkIt(sSQL))

  redim arrData(2)
  arrData(0) = "" & strPicTablePrefix & "PIC_CATEGORIES"
  arrData(1) = "CAT_NAME"
  arrData(2) = "'SkyPortal'"
  populateB(arrData)
end sub

sub crPicSubCatTbl()
  response.Write("<hr><b>Create PIC_SUBCATEGORIES table</b><br><br>")
  '::::::::::::::::::::: CREATE PIC_SUBCATEGORIES  TABLE ::::::::::::::::::::::::::
  sSQL = "CREATE TABLE [" & strPicTablePrefix & "PIC_SUBCATEGORIES]([CAT_ID] LONG, [SUBCAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [SUBCAT_IMAGE] TEXT(100), [SUBCAT_NAME] TEXT(100), [SUBCAT_THUMB] TEXT(100),[C_ORDER] INT DEFAULT 1,[GROUPS] MEMO);"

  createTable(checkIt(sSQL))

  redim indexes(0)
  indexes(0) = "CREATE INDEX [CAT_ID] ON [" & strPicTablePrefix & "PIC_SUBCATEGORIES]([CAT_ID]);"
  createIndx(indexes)

  '-------------------- populate table with default values --------------------------
  redim arrData(2)
  arrData(0) = "" & strPicTablePrefix & "PIC_SUBCATEGORIES "
  arrData(1) = "SUBCAT_NAME, CAT_ID"
  arrData(2) = "'Logo', 1"
  populateB(arrData)
end sub

sub crPicMainTbl()
  response.Write("<hr><b>Create PIC table</b><br><br>")
'::::::::::::::::::::::: CREATE PIC  TABLE :::::::::::::::::::::::::::
  'response.Write("<hr>Create PIC table<br><br>")
sSQL = "CREATE TABLE [" & strPicTablePrefix & "PIC]([BADLINK] LONG DEFAULT 0, [CATEGORY] LONG, [COPYRIGHT] TEXT(100), [DESCRIPTION] TEXT(255), [HIT] LONG DEFAULT 0, [KEYWORD] TEXT(255), [OWNER] TEXT(255), [PARENT_ID] LONG, [PIC_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [POST_DATE] TEXT(50), [POSTER] TEXT(100), [RATING] LONG NOT NULL DEFAULT 0, [SHOW] LONG DEFAULT 0, [TITLE] TEXT(100), [TURL] TEXT(255), [URL] TEXT(255), [VOTES] LONG NOT NULL DEFAULT 0, [FEATURED] BIT DEFAULT 0);"

createTable(checkIt(sSQL))

redim indexes(1)
indexes(0) = "CREATE INDEX [KEYWORD] ON [" & strPicTablePrefix & "PIC]([KEYWORD]);"
indexes(1) = "CREATE INDEX [PARENT_ID] ON [" & strPicTablePrefix & "PIC]([PARENT_ID]);"
createIndx(indexes)

'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO " & strPicTablePrefix & "PIC "
		strSql = strSql & "(TITLE"
		strSql = strSql & ", POST_DATE"
		strSql = strSql & ", DESCRIPTION"
		strSql = strSql & ", HIT"
		strSql = strSql & ", COPYRIGHT"
		strSql = strSql & ", POSTER"
		strSql = strSql & ", KEYWORD"
		strSql = strSql & ", CATEGORY" 
		strSql = strSql & ", PARENT_ID"
		strSql = strSql & ", SHOW"
		strSql = strSql & ", RATING"
		strSql = strSql & ", VOTES"
		strSql = strSql & ", OWNER"
		strSql = strSql & ", URL"
		strSql = strSql & ", TURL"
		strSql = strSql & ", BADLINK" 
		strSql = strSql & ", FEATURED" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'SkyPortal Logo'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 'SkyPortal logo example'"
		strSql = strSql & ", 1"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 'SkyDogg'"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '0'"
		strSql = strSql & ", 'http://www.skyportal.net/files/gallery_images/skyportal_logo_rs.jpg'"
		strSql = strSql & ", 'http://www.skyportal.net/files/gallery_images/skyportal_logo_sm.jpg'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub crPicRatingTbl()
  response.Write("<hr><b>Create PIC_RATING table</b><br><br>")
  ':::::::::::::::::::::::::: CREATE PIC_RATING  TABLE :::::::::::::::::::::::::::
  sSQL = "CREATE TABLE [" & strPicTablePrefix & "PIC_RATING]([COMMENTS] MEMO, [PIC] LONG, [RATE_BY] LONG, [RATE_DATE] TEXT(50), [RATING] LONG, [RATING_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL);"

  createTable(checkIt(sSQL))

  createIndex("CREATE INDEX [RATING_PICID] ON [" & strPicTablePrefix & "PIC_RATING]([PIC]);")
end sub

sub chkAppVerFld(typ)
	on error resume next
	Err.Clear
    sSql = "SELECT APP_VERSION FROM PORTAL_APPS WHERE APP_INAME = 'pictures'"
    set rsC = my_Conn.execute(sSql)
	if err.number <> 0 then
	  Err.Clear
	  strSql = "ALTER TABLE " & strTablePrefix & "APPS ADD APP_VERSION TEXT(10)"
	  my_Conn.Execute(checkIt(strSql)),,adCmdText + adExecuteNoRecords
	  if err.number <> 0 then
		Response.Write("<font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
	  else
		Response.Write("<b>" & txtDBTableAltSucc & "</b><br /><br />" & vbNewLine)
	  end if
	  Err.Clear
	end if
    set rsC = nothing
	if typ = 1 then
	  strSql = "UPDATE " & strTablePrefix & "APPS SET APP_VERSION='" & app_version & "' WHERE APP_INAME = 'pictures'"
	  my_Conn.Execute(checkIt(strSql))
	  Err.Clear
	end if
	on error goto 0
end sub

sub uninstall_Pictures()
  response.Write("<hr><b>Uninstall App</b><br><br>")
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'pictures'"
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
  
  sSql = "DELETE FROM "& strTablePrefix &"MODS WHERE M_NAME='"& apid &"'"
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "PAGES WHERE P_APP=" & apid
  executeThis(sSql)
 
  sSql = "delete from menu where APP_ID = " & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "UPLOAD_CONFIG WHERE UP_APPID=" & apid
  executeThis(sSql)
  
  droptable("" & strPicTablePrefix & "PIC_CATEGORIES")
  droptable("" & strPicTablePrefix & "PIC_SUBCATEGORIES")
  droptable("" & strPicTablePrefix & "PIC")
  droptable("" & strPicTablePrefix & "PIC_RATING")
	
  sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_INAME='pictures'"
  executeThis(sSql)
  mnu.DelMenuFiles("")
  response.Write("<b>Module Uninstall Complete</b><br><hr><br>")
end sub

':: start MODULE MENUS :::::::::::::::::::::::::::::::::::::::::::::::::
sub pictures_main_button()
  response.Write("<hr><h3>Pictures Module Menus</h3>")

  ct = 3
  response.Write("<hr><h4>Pictures Menu</h4><br>")
  mTitle = "Pictures"		' Friendly menu name
  mINAME = "m_pictures"		' Internal menu name = app_INAME : Must me different from mName
  mName = "Pictures"		' Link Head Text
  msName = "Picture"
  mCntFunct = "cntNewPictures()"
  mLink1 = "pic.asp"		'Main Directory
  mLink2 = "pic.asp?cmd=3"	'New
  mLink3 = "pic.asp?cmd=4"	'Popular
  mLink4 = "pic.asp?cmd=5"	'Top
  mLink5 = "pic.asp?cmd=8"	'Submit
  mLink6 = "openWindow3(''pic_pop.asp?mode=13'')"	'FAQ
 
  sSql = "delete from menu where APP_ID = " & app_id
  executeThis(sSql)

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuFunction,mnuTitle,INAME,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'"& mName &"', '"& mINAME &"','','','','','" & mCntFunct & "','"& mTitle &"','"& mINAME &"',"& app_id &",'1,2,3',1"
  populateB(arrData)

  sSql = "select ID from menu where Name = '"& mName &"' and INAME = '"& mINAME &"'"
  set rsT = my_Conn.execute(sSql)
  pID = rsT(0)
  set rsT = nothing

  redim arrData(7)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,mnuImage,onClick,mnuFunction,mnuTitle,INAME,ParentID,app_id,mnuAccess,mnuOrder"
  arrData(2) = "'Main Directory', '" & mName & "','" & mLink1 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2,3',1"
  arrData(3) = "'New "& mName &"', '" & mName & "','" & mLink2 & "','_parent','','','" & mCntFunct & "','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2,3',2"
  arrData(4) = "'Popular "& mName &"', '" & mName & "','" & mLink3 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',3"
  arrData(5) = "'Top "& mName &"', '" & mName & "','" & mLink4 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2,3',4"
  arrData(6) = "'Submit "& msName &"', '" & mName & "','" & mLink5 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2',5"
  arrData(7) = "'"& mName &" FAQ', '" & mName & "','','_blank','','" & mLink6 & "','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2,3',6"
  populateB(arrData)
 
 ':: add 'nav_main' menu reference
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuAccess,mnuOrder"
  arrData(2) = "'" & mTitle & " Menu', 'nav_main','','','','','Portal Navbar','nav_main',"& app_id &",'" & mINAME & "','1,2,3',3"
  populateB(arrData)

 ':: add 'Main Default' menu REFERENCE
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuAccess,mnuOrder"
  arrData(2) = "'" & mTitle & " Menu', 'def_main','','','','','Portal Default','def_main',"& app_id &",'" & mINAME & "','1,2,3',3"
  populateB(arrData)

end sub
':: end MODULE MAIN menus :::::::::::::::::::::::::::::::::::::::::


':: start MODULE ADMIN MENU :::::::::::::::::::::::::::::::::::::::::::::
sub pictures_admin_button()
ct = 3
  response.Write("<hr><h4>Pictures SAdmin Menu</h4><br>")
  mINAME = "pictures_admin"
  mTitle = "* Pictures ADMIN"
  mName = "Pictures"
  msName = "Picture"
  mLink1 = "admin_pic_admin.asp"
  mLink2 = "admin_pic_main.asp"
 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuOrder"
arrData(2) = "'" & mName & "', '" & mINAME & "','','','','','" & mTitle & "','" & mINAME & "',"& app_id &",1"
populateB(arrData)

sSql = "select ID from menu where Name = '"& mName &"' and INAME = '"& mINAME &"'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(12)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,mnuImage,onClick,mnuTitle,INAME,ParentID,app_id,mnuOrder"

arrData(2) = "'Approve New', '" & mName & "','" & mLink2 & "','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",1"
arrData(3) = "'Bad Links', '" & mName & "','" & mLink1 & "?cmd=40','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",2"
arrData(4) = "'Create Category', '" & mName & "','" & mLink1 & "','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",3"
arrData(5) = "'Edit Category', '" & mName & "','" & mLink1 & "?cmd=2','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",4"
arrData(6) = "'Delete Category', '" & mName & "','" & mLink1 & "?cmd=4','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",5"
arrData(7) = "'Create SubCategory', '" & mName & "','" & mLink1 & "?cmd=1','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",6"
arrData(8) = "'Edit SubCategory', '" & mName & "','" & mLink1 & "?cmd=5','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",7"
arrData(9) = "'Delete SubCategory', '" & mName & "','" & mLink1 & "?cmd=8','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",8"
arrData(10) = "'Edit " & msName & "', '" & mName & "','" & mLink1 & "?cmd=10','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",9"
arrData(11) = "'Delete " & msName & "', '" & mName & "','" & mLink1 & "?cmd=20','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",10"
arrData(12) = "'Browse " & mName & "', '" & mName & "','" & mLink1 & "?cmd=30','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",11"

populateB(arrData)
 
 ':: add module links to module admin menu
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'm_admin','','','','','Module Admin','m_admin',"& app_id &",'" & mINAME & "',1"
populateB(arrData)
 
 ':: add module links to superadmin menu
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'sadmin','','','','','Portal Admin','sadmin',"& app_id &",'" & mINAME & "',4"
'populateB(arrData)

end sub
':: end MODULE ADMIN menu ::::::::::::::::::::::::::::::::::::::::::::::::

sub pic_Upgrade09()
  response.Write("<hr><hr><b>UPDATE " & strPicTablePrefix & "PICTURES MODULE</b><br>")
  response.Write("<hr><b>Update " & strTablePrefix & "APPS table</b><br><br>")
	   strSql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_USERS = '1,2,3'"
	   strSql = strSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
	   strSql = strSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & DateToStr(now()) & "'"
	   strSql = strSql & ", APP_SUBSEC = 0 WHERE APP_INAME = '"& strPicTablePrefix &"pictures';"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update "& strPicTablePrefix &"PIC_CATEGORIES table</b><br><br>")
	   strSql = ""& strPicTablePrefix &"PIC_CATEGORIES"
	   strSql = strSql & ",[CG_READ] MEMO NULL,[CG_WRITE] MEMO NULL,[CG_FULL] MEMO NULL,[CG_INHERIT] INT NULL DEFAULT 1,[CG_PROPAGATE] INT NULL DEFAULT 1"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE " & strPicTablePrefix & "PIC_CATEGORIES "
	   strSql = strSql & "SET CG_READ = '1,2,3', CG_WRITE = '1,2', CG_FULL = '1', CG_INHERIT = 1 "
	   strSql = strSql & "WHERE CAT_ID not like 0;"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update "& strPicTablePrefix &"PIC_SUBCATEGORIES table</b><br><br>")
	   strSql = ""& strPicTablePrefix &"PIC_SUBCATEGORIES"
	   strSql = strSql & ",[SG_READ] MEMO NULL,[SG_WRITE] MEMO NULL,[SG_FULL] MEMO NULL,[SG_INHERIT] INT NULL"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE " & strPicTablePrefix & "PIC_SUBCATEGORIES "
	   strSql = strSql & "SET SG_READ = '1,2,3', SG_WRITE = '1,2', SG_FULL = '1', SG_INHERIT = 1 "
	   strSql = strSql & "WHERE SUBCAT_ID not like 0;"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update " & strTablePrefix & "UPLOAD_CONFIG table</b><br><br>")
	   strSql= "UPDATE " & strTablePrefix & "UPLOAD_CONFIG SET UP_ALLOWEDGROUPS = '1,2' WHERE UP_APPID = " & app_id & ";"
	   executeThis(strSql)
	
  response.Write("<hr><b>Update " & strTablePrefix & "FP table</b><br><br>")
	   strSql= "UPDATE " & strTablePrefix & "FP SET fp_groups = '1,2,3' WHERE APP_ID = " & app_id & ";"
	   executeThis(strSql)
	   
	redim indexes(3)
	indexes(0) = "CREATE INDEX [PSUB] ON [" & strPicTablePrefix & "PIC]([CATEGORY]);"
	indexes(1) = "CREATE INDEX [PUID] ON [" & strPicTablePrefix & "PIC]([PIC_ID]);"
	indexes(2) = "CREATE INDEX [SCAT_ID] ON [" & strPicTablePrefix & "PIC_SUBCATEGORIES]([SUBCAT_ID]);"
	indexes(3) = "CREATE INDEX [CAT_ID] ON [" & strPicTablePrefix & "PIC_CATEGORIES]([CAT_ID]);"
	createIndx(indexes)
	
	':: create new menu buttons/links
	pictures_main_button()
	pictures_admin_button()
	mnu.DelMenuFiles("")
end sub

sub pictures_Upgrade_10_11()
  response.Write("<hr><hr><b>UPDATE Pictures Module From v0.10 to v0.11</b><br />")
  Response.Write("<br /><b>Add Fields to PIC table</b><br />")
  strSql = "ALTER TABLE PIC ADD [UPDATED] TEXT(50) DEFAULT 0"
  alterTable(checkIt(strSql))
  strSql = "UPDATE PIC SET UPDATED = 0;"
  executeThis(strSql)
	
  strSql = "ALTER TABLE PIC ADD [ACTIVE] INT"
  alterTable(checkIt(strSql))
  strSql = "UPDATE PIC SET ACTIVE = 1 WHERE SHOW = 1;"
  executeThis(strSql)
  strSql = "UPDATE PIC SET ACTIVE = 0 WHERE SHOW = 0;"
  executeThis(strSql)
  strSql = "ALTER TABLE PIC DROP COLUMN [SHOW];" 
  alterTable(checkIt(strSql))
  
  strSql = "ALTER TABLE PIC ADD [TDATA1] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA2] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA3] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA4] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA5] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA6] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA7] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA8] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA9] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE PIC ADD [TDATA10] TEXT(255) NULL"
  alterTable(checkIt(strSql))

  Response.Write("<br /><b>Updated PIC table</b><br />")
  addMODs()
  updateVersion app_version,"pictures"
end sub

sub addMODs()
	redim arrData(4)
	arrData(0) = strTablePrefix & "MODS"
	arrData(1) = "M_NAME, M_CODE, M_VALUE"
	arrData(2) = "'" & app_id & "', 'admTaskLnk', 'pictures_adminPndLink()'"
	arrData(3) = "'" & app_id & "','siteSrch','pictures_SiteSearch()'"
	arrData(4) = "'" & app_id & "', 'pndTskCnt', 'pictures_PendTaskCnt()'"
	populateB(arrData)
end sub
%>
