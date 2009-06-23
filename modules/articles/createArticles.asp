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
dim txtIpsumSum, txtIpsum, txtIpsumSum2, txtIpsum2
bUninstall = false
bReinstall = false

app_version = "0.10"
do_app = true
incArtFp = false
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
  if incArtFp then
    createArticles()
	mnu.DelMenuFiles("")
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p><b>")
    Response.Write("You must add the fp_article.asp ""include"" file to<br>")
    Response.Write("your fp_custom.asp file in order ")
    Response.Write("to install this module</b></p><p>&nbsp;</p>")
    spThemeBlock1_close(intSkin)
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
    Response.Write("<p>&nbsp;</p>")
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

%>	</td>
    <td class="rightPgCol">
	<% intSkin = getSkin(intSubSkin,3) %>
	</td>
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" --><%

sub article_Upgrades()
  article_Upgrade08()
  migrateIntegratedTables("ARTICLE")
  article_Upgrade09_10()
end sub

sub createArticles()
txtIpsumSum = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Cras tempus orci a magna malesuada pharetra. Mauris pede dolor, varius at, consectetuer aliquet, sagittis eget, metus. Nunc sit amet pede."
txtIpsum = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Cras tempus orci a magna malesuada pharetra. Mauris pede dolor, varius at, consectetuer aliquet, sagittis eget, metus. Nunc sit amet pede. Proin eu turpis. Sed ipsum mi, condimentum ut, ullamcorper a, mattis eget, turpis. Pellentesque ipsum. Quisque iaculis risus non elit. Nullam ut velit. Cras semper risus sit amet dolor. Morbi semper nunc et odio. Aliquam adipiscing tortor quis eros. Pellentesque ac lacus. Mauris blandit. Phasellus sed purus. Integer elit. Donec ut magna vel diam interdum imperdiet. Curabitur non pede viverra sem vehicula congue. Nullam elementum enim at ipsum. Curabitur ac quam.<br><br>Quisque a risus quis est pulvinar sodales. Aenean sollicitudin. Donec imperdiet odio id neque. Curabitur tellus. Pellentesque vitae est. Cras sit amet massa sit amet libero gravida condimentum. Donec id risus. Morbi tempor condimentum velit. Nunc sodales diam sit amet enim. Donec eleifend massa eget felis. Nunc elit.<br><br>Ut tellus wisi, convallis vel, dapibus quis, tincidunt in, tellus. Integer convallis purus. Fusce at mi sit amet orci aliquet imperdiet. Sed nec eros. Sed vel justo sit amet dolor nonummy lobortis. Nunc dictum dolor ac turpis. Sed risus ante, pulvinar et, tempus a, malesuada sit amet, libero. Aliquam ultrices. Integer consectetuer, libero sed auctor tempus, turpis neque venenatis nisl, eu luctus purus lacus et arcu. Integer dapibus, justo nec ultrices tincidunt, erat wisi convallis tortor, vel scelerisque nulla turpis eu tortor. Nunc et ipsum. Ut eu risus. Mauris ultrices augue in est. Sed id est a quam gravida varius. Aliquam fringilla, dui a vulputate cursus, ligula ipsum vehicula eros, vitae tempus sapien odio quis tellus."

txtIpsumSum2 = "Morbi tempor sagittis nibh. Maecenas nulla. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Quisque quis odio. Sed orci velit, laoreet non, commodo a, sollicitudin ut, tortor."
txtIpsum2 = "Nunc et justo. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ut sollicitudin, magna a pellentesque imperdiet, turpis mauris dapibus ipsum, nec placerat nibh tortor et felis. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Aliquam id est ut magna varius cursus. Duis adipiscing. Nulla vel sem sed leo vulputate pharetra. Pellentesque ante augue, tincidunt at, facilisis id, interdum ac, sem. Vivamus tristique, augue eu venenatis lobortis, est massa porta elit, vitae blandit tellus lacus quis orci. Morbi feugiat nulla nec eros tincidunt venenatis. Nullam elementum, libero eget rutrum malesuada, orci purus lobortis tellus, vel pellentesque felis lorem vitae ligula. Sed id augue. Nunc orci libero, dapibus hendrerit, auctor id, vestibulum id, est. Praesent venenatis, dolor eget aliquet accumsan, nunc nisi tincidunt ligula, vitae consectetuer arcu lectus commodo arcu. Maecenas vel dolor. Vivamus in lorem eget quam tristique eleifend. Vestibulum elementum commodo turpis. Nullam sed nisi."

  spThemeBlock1_open(intSkin)
  response.Write("<hr><b>ARTICLES MODULE</b><br><br>")
  
  'check if app is existing
  sSql = "SELECT APP_NAME, APP_ID, APP_VERSION FROM PORTAL_APPS WHERE APP_INAME = 'article'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall or bReinstall then
      uninstall_Articles()
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing
  'do_app = true
Response.Write(cur_appVer)
 if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "0.10"
	   ':: new version
     case "0.9"
	   ':: existing version
	   article_Upgrade09_10()
	   migrateIntegratedTables("ARTICLE")
	   updateVersion app_version,"article"
     case "0.8"
	   article_Upgrade09_10()
	   migrateIntegratedTables("ARTICLE")
	   updateVersion app_version,"article"
	 case else
	   article_Upgrades()
	   updateVersion app_version,"article"
   end select
 elseif not bUninstall then
 
    '::::: Create APP ::::::::::::::::::::::::::::
   cr_App()
   
   crArtCatTbl()
   crArtSubCatTbl()
   crArtMainTbl()
   crArtRatingTbl()
   
   addArtFPitems()
   'addArtUploads()

   article_Upgrades()
 end if
 if not bUninstall then
  response.Write("<hr><h3>Articles Module Installed</h3><br><br>")
 else
  response.Write("<hr><h3>Articles Module Uninstalled</h3><br><br>")
 end if
  response.Write("<b>Be sure to delete this file createArticles.asp from your server!</b><br><br>")
  response.Write("<a href=""article.asp""><b>Continue</b></a><br><br><br><br>")
  spThemeBlock1_close(intSkin)
  Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
end sub

sub cr_App()
  response.Write("<hr><b>Update PORTAL_APPS</b><br><br>")
  'create the app
  redim arrData(2)
  arrData(0) = "[PORTAL_APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW],[APP_VERSION]"
  arrData(2) = "'article','article',1,0,'1,2,3',1,1,'config_articles','article.asp','" & app_version & "'"
  populateB(arrData)
  
  'return app_id
  sSql = "SELECT APP_ID FROM PORTAL_APPS WHERE APP_INAME = 'article'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub addArtFPitems()
	'::::::::::::::::::::: CREATE ARTICLE FRONT PAGE ITEMS :::::::::::::::::::::::::
  response.Write("<hr><b>Update PORTAL_FP table with ARTICLE info</b><br><br>")
	'articles front page
	redim arrData(9)
	arrData(0) = "[PORTAL_FP]"
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
end sub

sub crArtCatTbl()
  response.Write("<hr><b>Create ARTICLE_CATEGORIES table</b><br><br>")
	'::::::::::::::::::::::: CREATE ARTICLE_CATEGORIES  TABLE :::::::::::::::::::::::::
	sSQL = "CREATE TABLE [ARTICLE_CATEGORIES]([CAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CAT_NAME] TEXT(100),[C_ORDER] INT DEFAULT 1,[GROUPS] MEMO DEFAULT NULL)"

	createTable(checkIt(sSQL))

	redim arrData(6)
	arrData(0) = "ARTICLE_CATEGORIES"
	arrData(1) = "CAT_NAME, C_ORDER"
	arrData(2) = "'Entertainment',1"
	arrData(3) = "'Software Development',4"
	arrData(4) = "'Others',5"
	arrData(5) = "'Humor',2"
	arrData(6) = "'Internet',3"
	populateB(arrData)
end sub

sub crArtSubCatTbl()
  response.Write("<hr><b>Create ARTICLE_SUBCATEGORIES table</b><br><br>")
	'::::::::::::::::::::: CREATE ARTICLE_SUBCATEGORIES  TABLE :::::::::::::::::::::::::
	sSQL = "CREATE TABLE [ARTICLE_SUBCATEGORIES]([CAT_ID] LONG, [SUBCAT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [SUBCAT_NAME] TEXT(100), [C_ORDER] INT DEFAULT 1, [GROUPS] MEMO NULL);"

	createTable(checkIt(sSQL))

	redim indexes(0)
	indexes(0) = "CREATE INDEX [CAT_ID] ON [ARTICLE_SUBCATEGORIES]([CAT_ID]);"
	createIndx(indexes)

'-------------------- populate table with default values --------------------------
	redim arrData(8)
	arrData(0) = "ARTICLE_SUBCATEGORIES "
	arrData(1) = "SUBCAT_NAME, CAT_ID, C_ORDER"
	arrData(2) = "'Movies', 1, 1"
	arrData(3) = "'Others', 2, 1"
	arrData(4) = "'Temp', 3, 1"
	arrData(5) = "'Jokes', 4, 1"
	arrData(6) = "'Stories', 4, 2"
	arrData(7) = "'Others', 5, 1"
	arrData(8) = "'Others', 1, 2"
	populateB(arrData)
end sub

sub crArtMainTbl()
  response.Write("<hr><b>Create ARTICLE table</b><br><br>")
	'::::::::::::::::::::::::: CREATE ARTICLE TABLE ::::::::::::::::::::::::::::
	sSQL = "CREATE TABLE [ARTICLE]([ARTICLE_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [AUTHOR] TEXT(100), [AUTHOR_EMAIL] TEXT(100), [CATEGORY] LONG, [CONTENT] MEMO, [HIT] LONG DEFAULT 0, [KEYWORD] TEXT(255), [PARENT_ID] LONG, [POST_DATE] TEXT(50), [POSTER] TEXT(100), [POSTER_EMAIL] TEXT(100), [RATING] LONG NOT NULL DEFAULT 0, [SHOW] LONG, [SUMMARY] MEMO, [TITLE] TEXT(100), [VOTES] LONG NOT NULL DEFAULT 0, [FEATURED] INT DEFAULT 0)"

	createTable(checkIt(sSQL))

	redim indexes(3)
	indexes(0) = "CREATE INDEX [ARTICLE_ID] ON [ARTICLE]([ARTICLE_ID]);"
	indexes(1) = "CREATE INDEX [CATEGORY] ON [ARTICLE]([CATEGORY]);"
	indexes(2) = "CREATE INDEX [KEYWORD] ON [ARTICLE]([KEYWORD]);"
	indexes(3) = "CREATE INDEX [PARENT_ID] ON [ARTICLE]([PARENT_ID]);"
	createIndx(indexes)

	'-------------------- populate table with default values --------------------------
		strSql = "INSERT INTO ARTICLE "
		strSql = strSql & "(TITLE"
		strSql = strSql & ", POST_DATE"
		strSql = strSql & ", CONTENT"
		strSql = strSql & ", HIT"
		strSql = strSql & ", AUTHOR"
		strSql = strSql & ", AUTHOR_EMAIL"
		strSql = strSql & ", POSTER"
		strSql = strSql & ", POSTER_EMAIL"
		strSql = strSql & ", KEYWORD"
		strSql = strSql & ", CATEGORY"
		strSql = strSql & ", PARENT_ID"
		strSql = strSql & ", SHOW"
		strSql = strSql & ", RATING"
		strSql = strSql & ", VOTES"
		strSql = strSql & ", SUMMARY" 
		strSql = strSql & ", FEATURED" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'Lorem Ipsum 2'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & txtIpsum2 & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 'Anonymous'"
		strSql = strSql & ", ' '"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 2"
		strSql = strSql & ", 2"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & txtIpsumSum2 & "'"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
		
		strSql = "INSERT INTO ARTICLE "
		strSql = strSql & "(TITLE"
		strSql = strSql & ", POST_DATE"
		strSql = strSql & ", CONTENT"
		strSql = strSql & ", HIT"
		strSql = strSql & ", AUTHOR"
		strSql = strSql & ", AUTHOR_EMAIL"
		strSql = strSql & ", POSTER"
		strSql = strSql & ", POSTER_EMAIL"
		strSql = strSql & ", KEYWORD"
		strSql = strSql & ", CATEGORY"
		strSql = strSql & ", PARENT_ID"
		strSql = strSql & ", SHOW"
		strSql = strSql & ", RATING"
		strSql = strSql & ", VOTES"
		strSql = strSql & ", SUMMARY" 
		strSql = strSql & ", FEATURED" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'Lorem Ipsum'"
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & txtIpsum & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'www.ipsum.com'"
		strSql = strSql & ", 'http://www.lipsum.com/'"
		strSql = strSql & ", 'SkyDogg'"
		strSql = strSql & ", ' '"
		strSql = strSql & ", 'Lorem ipsum'"
		strSql = strSql & ", 7"
		strSql = strSql & ", 1"
		strSql = strSql & ", 1"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & txtIpsumSum & "'"
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub crArtRatingTbl()
  response.Write("<hr><b>Create ARTICLE_RATING table</b><br><br>")
	'::::::::::::::::::::: CREATE ARTICLE_RATING  TABLE :::::::::::::::::::::::::::::
	sSQL = "CREATE TABLE [ARTICLE_RATING]([ARTICLE] LONG, [COMMENTS] MEMO, [RATE_BY] LONG, [RATE_DATE] TEXT(50), [RATING] LONG, [RATING_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL)"

	createTable(checkIt(sSQL))
end sub

sub addArtUploads()
  response.Write("<hr><b>Update " & strTablePrefix & "UPLOAD_CONFIG table with ARTICLE info</b><br><br>")
		strSql = "INSERT INTO " & strTablePrefix & "UPLOAD_CONFIG "
		strSql = strSql & "(UP_SIZELIMIT"
		strSql = strSql & ", UP_ALLOWEDEXT"
		strSql = strSql & ", UP_LOGUSERS"
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
		strSql = strSql & ") VALUES ("
		strSql = strSql & "1000"
		strSql = strSql & ", 'gif,jpg,zip,rar'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 'article'"
		strSql = strSql & ", 1"
		strSql = strSql & ", 'upload.log'"
		strSql = strSql & ", " & app_id & ""
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"						
		strSql = strSql & ")"
'		response.Write(strSql)
		'populateA(strSql)
end sub

sub uninstall_Articles()
  response.Write("<hr><b>Uninstall App</b><br><br>")
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'article'"
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
	
  sSql = "DELETE FROM " & strTablePrefix & "M_PARENT WHERE APP_ID=" & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID=" & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE APP_ID=" & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "M_RATING WHERE APP_ID=" & apid
  executeThis(sSql)
	
  sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_INAME='article'"
  executeThis(sSql)
  
  droptable("ARTICLE")
  'droptable("ARTICLE_CATEGORIES")
  'droptable("ARTICLE_SUBCATEGORIES")
  'droptable("ARTICLE_RATING")
  
  response.Write("<b>Module Uninstall Complete</b><br><hr><br>")
end sub

'migrateIntegratedTables("ARTICLE")
sub migrateIntegratedTables(t)
  response.Write("<hr><hr><b>Migrate Integrated Module Tables</b>")
  response.Write("<br/><br/>")
  response.Write("<b>Migrate Integrated Categories and Subcats</b>")
  response.Write("<br /><br />")
  strSql = "ALTER TABLE " & t & " ADD [INTEGRATED] INT;"
  alterTable(checkIt(strSql))
  sSql = "UPDATE " & t & " SET INTEGRATED = 0"
  executeThis(sSql)
  
  sSql = "SELECT * FROM " & t & "_CATEGORIES ORDER BY CAT_ID"
  set rsC = my_Conn.execute(sSql)
  if not rsC.eof then
    do until rsC.eof
	  oldCatID = rsC("CAT_ID")
	  newCatID = integrateMCategory(rsC)
	  
  	  sSql = "SELECT * FROM " & t & "_SUBCATEGORIES WHERE CAT_ID = " & oldCatID & " ORDER BY SUBCAT_ID"
  	  set rsS = my_Conn.execute(sSql)
  	  if not rsS.eof then
    	do until rsS.eof
	  	  oldSCatID = rsS("SUBCAT_ID")
	      newSCatID = integrateMSubCategory(rsS,newCatID)
	  
  	  	  sSql = "UPDATE " & t & " SET CATEGORY="& newSCatID
  	  	  sSql = sSql & ", INTEGRATED = 1"
  	  	  sSql = sSql & " WHERE CATEGORY = " & oldSCatID
  	  	  sSql = sSql & " AND INTEGRATED = 0"
  	  	  executeThis(sSql)
	  	  rsS.movenext
	  	loop
  	  end if
  	  set rsS = nothing
	  
	  rsC.movenext
	loop
  end if
  set rsC = nothing
  
  response.Write("<b>Migrate Integrated Ratings/Comments</b><br /><br />")
  sSql = "SELECT * FROM " & t & "_RATING ORDER BY RATING_ID"
  set rsR = my_Conn.execute(sSql)
  if not rsR.eof then
    do until rsR.eof
  	  integrateMRatings rsR,t
	  rsR.movenext
	loop
  end if
  set rsR = nothing
  
  ':: remove orphaned articles
  sSql = "DELETE FROM " & t & " WHERE INTEGRATED = 0"
  executeThis(sSql)
  
  ':: clean up
  strSql = "ALTER TABLE " & t & " DROP COLUMN [INTEGRATED];" 
  alterTable(checkIt(strSql))
  
  droptable("" & t & "_CATEGORIES")
  droptable("" & t & "_SUBCATEGORIES")
  droptable("" & t & "_RATING")
end sub

sub integrateMRatings(obj,t)
  sSql = "INSERT INTO " & strTablePrefix & "M_RATING ("
  sSQL = sSql & "ITEM_ID,RATE_BY,RATE_DATE"
  sSQL = sSql & ",RATING,COMMENTS,APP_ID"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & obj(t)
  sSQL = sSql & "," & obj("RATE_BY")
  sSQL = sSql & ",'" & obj("RATE_DATE") & "'"
  sSQL = sSql & "," & obj("RATING")
  sSQL = sSql & ",'" & replace(obj("COMMENTS"),"'","") & "'"
  sSQL = sSql & "," & app_id & ""
  sSQL = sSql & ")"
  executeThis(sSQL)
end sub

function integrateMSubCategory(obj,cat)
  tCount = getCount("ARTICLE_ID","ARTICLE","CATEGORY=" & obj("SUBCAT_ID") & "")
  
  sSql = "INSERT INTO " & strTablePrefix & "M_SUBCATEGORIES ("
  sSQL = sSql & "SUBCAT_NAME,SUBCAT_SDESC,SUBCAT_LDESC"
  sSQL = sSql & ",CAT_ID,SG_READ,SG_WRITE,SG_FULL"
  sSQL = sSql & ",SG_INHERIT,APP_ID,C_ORDER,ITEM_CNT"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'" & obj("SUBCAT_NAME") & "'"
  sSQL = sSql & ",''"
  sSQL = sSql & ",''"
  sSQL = sSql & "," & cat
  sSQL = sSql & ",'1,2,3'"
  sSQL = sSql & ",'1,2'"
  sSQL = sSql & ",'1'"
  sSQL = sSql & ",1"
  sSQL = sSql & "," & app_id
  sSQL = sSql & "," & obj("C_ORDER")
  sSQL = sSql & "," & tCount
  sSQL = sSql & ")"
  executeThis(sSQL)
  
  sSql = "SELECT SUBCAT_ID FROM " & strTablePrefix & "M_SUBCATEGORIES ORDER BY SUBCAT_ID DESC"
  set rsT = my_Conn.execute(sSql)
    t_id = rsT("SUBCAT_ID")
  set rsT = nothing
  
  integrateMSubCategory = t_id
end function

function integrateMCategory(obj)
  response.Write("<b>Migrate Integrated Categories</b><br /><br />")
  sSql = "INSERT INTO " & strTablePrefix & "M_CATEGORIES ("
  sSQL = sSql & "CAT_NAME,CAT_SDESC,CAT_LDESC"
  sSQL = sSql & ",CG_READ,CG_WRITE,CG_FULL,CG_INHERIT"
  sSQL = sSql & ",CG_PROPAGATE,APP_ID,C_ORDER"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'" & obj("CAT_NAME") & "'"
  sSQL = sSql & ",''"
  sSQL = sSql & ",''"
  sSQL = sSql & ",'1,2,3'"
  sSQL = sSql & ",'1,2'"
  sSQL = sSql & ",'1'"
  sSQL = sSql & ",1"
  if obj("CG_PROPAGATE") <> "" then
  sSQL = sSql & ",1"
  else
  sSQL = sSql & ",1"
  end if
  sSQL = sSql & "," & app_id & ""
  sSQL = sSql & "," & obj("C_ORDER") & ""
  sSQL = sSql & ")"
  executeThis(sSQL)
  
  sSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_CATEGORIES ORDER BY CAT_ID DESC"
  set rsT = my_Conn.execute(sSql)
    t_id = rsT("CAT_ID")
  set rsT = nothing
  
  integrateMCategory = t_id
end function

sub addSkypage()
  response.Write("<hr><b>Add data to PORTAL_FP table</b><br />")
	redim arrData(3)
	arrData(0) = "[PORTAL_FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Articles Menu','art_menu','menu_art',1,4,'Default Articles Manager menu.','1,2,3'," & app_id & ""
	arrData(3) = "'Articles Intro','art_intro','mod_displayIntro:" & app_id & "',1,4,'Default Articles Intro.','1,2,3'," & app_id & ""
	populateB(arrData)
	
  Response.Write("<br /><b>Create SkyPage</b><br />")
		strSql = "INSERT INTO " & strTablePrefix & "PAGES "
		strSql = strSql & "(P_NAME"
		strSql = strSql & ", P_INAME"
		strSql = strSql & ", P_TITLE"
		strSql = strSql & ", P_CONTENT"
		strSql = strSql & ", P_ACONTENT"
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
		strSql = strSql & "'Article Manager'"  'P_NAME
		strSql = strSql & ", 'article'" 	'P_INAME
		strSql = strSql & ", 'Article Manager'" 'P_TITLE
		strSql = strSql & ", ' '" 'P_CONTENT
		strSql = strSql & ", ' '" 'P_ACONTENT
		strSql = strSql & ", 'Article Menu:menu_art,Article - Newest:article_sm:new'" 'P_LEFTCOL
		strSql = strSql & ", 'Article - Featured:article_sm:featured,Article - Popular:article_sm:top,Article - Random:article_sm:random'" 'P_RIGHTCOL
		strSql = strSql & ", 'Article - Intro:mod_displayIntro:" & app_id & "'" 'P_MAINTOP
		strSql = strSql & ", 'Article - Newest:article_lg:new'" 'P_MAINBOTTOM
		strSql = strSql & ", " & app_id & ""  'P_APP
		strSql = strSql & ", 0"  'P_USE_PG_DISP
		strSql = strSql & ", 'article.asp'" 'P_OTHER_URL
		strSql = strSql & ", 0"  'P_CAN_DELETE
		strSql = strSql & ", ''" 'P_META_TITLE
		strSql = strSql & ", ''" 'P_META_DESC
		strSql = strSql & ", ''" 'P_META_KEY
		strSql = strSql & ", ''" 'P_META_EXPIRES
		strSql = strSql & ", ''" 'P_META_RATING
		strSql = strSql & ", ''" 'P_META_DIST
		strSql = strSql & ", ''" 'P_META_ROBOTS						
		strSql = strSql & ")"
		'response.Write(strSql)
		populateA(strSql)
	
  Response.Write("<br /><b>Create SkyPage</b><br />")
		strSql = "INSERT INTO " & strTablePrefix & "PAGES "
		strSql = strSql & "(P_NAME"
		strSql = strSql & ", P_INAME"
		strSql = strSql & ", P_TITLE"
		strSql = strSql & ", P_CONTENT"
		strSql = strSql & ", P_ACONTENT"
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
		strSql = strSql & "'Article Read'"  'P_NAME
		strSql = strSql & ", 'article_read'" 	'P_INAME
		strSql = strSql & ", 'Read Article'" 'P_TITLE
		strSql = strSql & ", ' '" 'P_CONTENT
		strSql = strSql & ", ' '" 'P_ACONTENT
		strSql = strSql & ", 'Article Menu:menu_art,Article - Newest:article_sm:new'" 'P_LEFTCOL
		strSql = strSql & ", ''" 'P_RIGHTCOL
		strSql = strSql & ", ''" 'P_MAINTOP
		strSql = strSql & ", ''" 'P_MAINBOTTOM
		strSql = strSql & ", " & app_id & ""  'P_APP
		strSql = strSql & ", 0"  'P_USE_PG_DISP
		strSql = strSql & ", 'article_read.asp'" 'P_OTHER_URL
		strSql = strSql & ", 0"  'P_CAN_DELETE
		strSql = strSql & ", ''" 'P_META_TITLE
		strSql = strSql & ", ''" 'P_META_DESC
		strSql = strSql & ", ''" 'P_META_KEY
		strSql = strSql & ", ''" 'P_META_EXPIRES
		strSql = strSql & ", ''" 'P_META_RATING
		strSql = strSql & ", ''" 'P_META_DIST
		strSql = strSql & ", ''" 'P_META_ROBOTS						
		strSql = strSql & ")"
		'response.Write(strSql)
		populateA(strSql)
end sub

sub addIntro()
  Response.Write("<br /><b>Add module intro</b><br />")
  sImsg = "<b>Welcome to the <span class=""fAlert""><b>NEW</b></span> SkyPortal <i>Articles Manager</i> Module.</b>"
  sImsg = sImsg & "<br/><br/>This is your module introduction block. You can create any message that you want. You can edit this message by clicking the *edit* icon in the title bar above this message."
  sImsg = sImsg & "<br/><br/>You can edit the module page layout through the admin Layout Manager. Any changes will affect all pages of this module."
  sImsg = sImsg & "<br/>"
  'sImsg = sImsg & "<br/>"
  
		strSql = "INSERT INTO " & strTablePrefix & "WELCOME "
		strSql = strSql & "(W_TITLE"
		strSql = strSql & ", W_SUBJECT"
		strSql = strSql & ", W_MESSAGE"
		strSql = strSql & ", W_DELETE"
		strSql = strSql & ", W_MODULE"
		strSql = strSql & ", W_ACTIVE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'Article Intro'"
		strSql = strSql & ", 'Articles Introduction'"
		strSql = strSql & ", '" & sImsg & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub newAdminMenu()
  sSql = "DELETE FROM menu where INAME = 'articles_admin'"
  executeThis(sSql)
  ct = 3
  mINAME = "articles_admin"
  mTitle = "* Articles ADMIN"
  mName = "Articles"
  mLink1 = "admin_articles.asp"
  
  response.Write("<hr><h4>" & mName & " SAdmin Menu</h4><br />")
 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuOrder"

arrData(2) = "'" & mName & "', '" & mINAME & "','','','','','" & mTitle & "','" & mINAME & "',"& app_id &",1"
populateB(arrData)

sSql = "select ID from menu where Name = '"& mName &"' and INAME = '"& mINAME &"'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(4)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,mnuImage,onClick,mnuTitle,INAME,ParentID,app_id,mnuOrder"

arrData(2) = "'Attention Items', '" & mName & "','" & mLink1 & "','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",1"
arrData(3) = "'Category Manager', '" & mName & "','" & mLink1 & "?cmd=20','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",2"
arrData(4) = "'Subcategory Manager', '" & mName & "','" & mLink1 & "?cmd=21','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",3"

populateB(arrData)

':: ADD to Module Admin menu 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,mnuAdd,mnuTitle,INAME,onClick,app_id,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'm_admin','articles_admin','Module Admin','m_admin','',"& app_id &",1"
populateB(arrData)
  
end sub

sub addMODs()
    Response.Write("<br /><b>Add Pending Tasks</b><br />")
    Response.Write("<br /><b>Add Site Search</b><br />")
	redim arrData(4)
	arrData(0) = strTablePrefix & "MODS"
	arrData(1) = "M_NAME, M_CODE, M_VALUE"
	arrData(2) = "'" & app_id & "','admTaskLnk','article_adminPndLink()'"
	arrData(3) = "'" & app_id & "','siteSrch','article_SiteSearch()'"
	arrData(4) = "'" & app_id & "', 'pndTskCnt', 'article_PendTaskCnt()'"
	populateB(arrData)
end sub

sub addTData(t)
  strSql = "ALTER TABLE " & t & " ADD [TDATA1] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA2] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA3] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA4] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA5] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA6] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA7] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA8] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA9] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE " & t & " ADD [TDATA10] TEXT(255) NULL"
  alterTable(checkIt(strSql))
end sub

sub article_Upgrade09_10()
  response.Write("<hr><hr><b>UPDATE ARTICLES MODULE From v0.09 to v0.10</b><br />")
  response.Write("<hr><b>Update " & strTablePrefix & "APPS table</b><br /><br />")
  strSql = "UPDATE "&strTablePrefix&"APPS SET APP_GROUPS_USERS = '1,2,3'"
  strSql = strSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
  strSql = strSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & DateToStr(now()) & "'"
  strSql = strSql & ", APP_SUBSEC = 1 WHERE APP_INAME = 'article';"
  executeThis(strSql)
  
  Response.Write("<br /><b>Add ARTICLE table Fields</b><br />")
  addTData("ARTICLE")
  
  strSql = "ALTER TABLE ARTICLE ADD [ACTIVE] INT"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE ARTICLE ADD [UPDATED] TEXT(50) DEFAULT 0 NULL"
  alterTable(checkIt(strSql))
  Response.Write("<br /><b>Update table fields</b><br />")
  strSql = "UPDATE ARTICLE SET ACTIVE = 1 WHERE SHOW = 1;"
  executeThis(strSql)
  strSql = "UPDATE ARTICLE SET ACTIVE = 0 WHERE SHOW = 0;"
  executeThis(strSql)
  strSql = "UPDATE ARTICLE SET UPDATED = '0';"
  executeThis(strSql)
  
  strSql = "UPDATE ARTICLE SET ARTICLE.TDATA2 = ARTICLE.AUTHOR WHERE ARTICLE_ID > 0"
  executeThis(strSql)
  strSql = "UPDATE ARTICLE SET ARTICLE.TDATA3 = ARTICLE.AUTHOR_EMAIL WHERE ARTICLE_ID > 0"
  executeThis(strSql)
  
  Response.Write("<br /><b>Drop columns</b><br />")
  strSql = "ALTER TABLE ARTICLE DROP COLUMN [SHOW];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE ARTICLE DROP COLUMN [AUTHOR];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE ARTICLE DROP COLUMN [AUTHOR_EMAIL];" 
  alterTable(checkIt(strSql))
  
  ':: redo featured column
  strSql = "ALTER TABLE ARTICLE DROP INDEX [DF__ARTICLE__FEATURE__7EF6D905];"
  'alterTable(checkIt(strSql))
  on error resume next 
  my_Conn.execute(strSql)
  on error goto 0
  strSql = "ALTER TABLE ARTICLE DROP COLUMN [FEATURED];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE ARTICLE ADD [FEATURED] INT"
  alterTable(checkIt(strSql))
  strSql = "UPDATE ARTICLE SET FEATURED = 0;"
  executeThis(strSql)
  
  addMODs()
  addSkypage()
  addIntro()
  newAdminMenu()
end sub

sub article_Upgrade08()
  response.Write("<hr><hr><b>UPDATE ARTICLE MODULE</b><br>")
  response.Write("<hr><b>Update " & strTablePrefix & "APPS table</b><br><br>")
	   strSql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_USERS = '1,2,3'"
	   strSql = strSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
	   strSql = strSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & DateToStr(now()) & "'"
	   strSql = strSql & ", APP_SUBSEC = 0 WHERE APP_INAME = 'article';"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update ARTICLE_CATEGORIES table</b><br><br>")
	   strSql = "ARTICLE_CATEGORIES"
	   strSql = strSql & ",[CG_READ] MEMO NULL,[CG_WRITE] MEMO NULL,[CG_FULL] MEMO NULL,[CG_INHERIT] INT NULL DEFAULT 1,[CG_PROPAGATE] INT NULL DEFAULT 1"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE ARTICLE_CATEGORIES "
	   strSql = strSql & "SET CG_READ = '1,2,3', CG_WRITE = '1,2', CG_FULL = '1', CG_INHERIT = 1 "
	   strSql = strSql & "WHERE CAT_ID not like 0;"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update ARTICLE_SUBCATEGORIES table</b><br><br>")
	   strSql = "ARTICLE_SUBCATEGORIES"
	   strSql = strSql & ",[SG_READ] MEMO NULL,[SG_WRITE] MEMO NULL,[SG_FULL] MEMO NULL,[SG_INHERIT] INT NULL"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE ARTICLE_SUBCATEGORIES "
	   strSql = strSql & "SET SG_READ = '1,2,3', SG_WRITE = '1,2', SG_FULL = '1', SG_INHERIT = 1 "
	   strSql = strSql & "WHERE SUBCAT_ID not like 0;"
	   executeThis(strSql)
	
  response.Write("<hr><b>Update " & strTablePrefix & "FP table</b><br><br>")
	   strSql= "UPDATE " & strTablePrefix & "FP SET fp_groups = '1,2,3' WHERE APP_ID = " & app_id & ";"
	   executeThis(strSql)
	   
	redim indexes(1)
	indexes(0) = "CREATE INDEX [A_SCAT_ID] ON [ARTICLE_SUBCATEGORIES]([SUBCAT_ID]);"
	indexes(1) = "CREATE INDEX [A_CAT_ID] ON [ARTICLE_CATEGORIES]([CAT_ID]);"
	createIndx(indexes)
	
	article_main_button()
	article_admin_button()
end sub


':: start MODULE MENUS :::::::::::::::::::::::::::::::::::::::::::::::::
sub article_main_button()
ct = 3
  mTitle = "Articles"		' Friendly menu name
  mINAME = "m_article"		' Internal menu name = app_INAME : Must me different from mName
  mName = "Articles"		' Link Head Text
  msName = "Article"
  mCntFunct = "cntNewArticles()"
  mLink1 = "article.asp"		'Main Directory
  mLink2 = "article.asp?cmd=3"	'New
  mLink3 = "article.asp?cmd=4"	'Popular
  mLink4 = "article.asp?cmd=5"	'Top
  mLink5 = "article.asp?cmd=7"	'Submit
  mLink6 = "openWindow3(''article_pop.asp?mode=10'')"	'FAQ
  response.Write("<hr><h4>" & mName & " Menu</h4><br>")
 
  sSql = "delete from menu where APP_ID = " & app_id
  executeThis(sSql)

redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuFunction,mnuTitle,INAME,app_id,mnuOrder"
arrData(2) = "'"& mName &"', '"& mINAME &"','','','','','" & mCntFunct & "','"& mTitle &"','"& mINAME &"',"& app_id &",1"
populateB(arrData)

sSql = "select ID from menu where Name = '"& mName &"' and INAME = '"& mINAME &"'"
set rsT = my_Conn.execute(sSql)
pID = rsT(0)
set rsT = nothing

redim arrData(7)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,mnuImage,onClick,mnuFunction,mnuTitle,INAME,ParentID,app_id,mnuAccess,mnuOrder"
arrData(2) = "'Main Directory', '" & mName & "','" & mLink1 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',1"
arrData(3) = "'New "& mName &"', '" & mName & "','" & mLink2 & "','_parent','','','" & mCntFunct & "','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',2"
arrData(4) = "'Popular "& mName &"', '" & mName & "','" & mLink3 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',3"
arrData(5) = "'Top "& mName &"', '" & mName & "','" & mLink4 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',4"
arrData(6) = "'Submit "& msName &"', '" & mName & "','" & mLink5 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2',5"
arrData(7) = "'"& mName &" FAQ', '" & mName & "','','_blank','','" & mLink6 & "','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'',6"
populateB(arrData)
 
 ':: add 'nav_main' menu reference
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'nav_main','','','','','Portal Navbar','nav_main',"& app_id &",'" & mINAME & "',3"
populateB(arrData)

 ':: add 'Main Default' menu REFERENCE
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuAccess,mnuOrder"
  arrData(2) = "'" & mTitle & " Menu', 'def_main','','','','','Portal Default','def_main',"& app_id &",'" & mINAME & "','',3"
  populateB(arrData)

end sub
':: end MODULE MAIN menus :::::::::::::::::::::::::::::::::::::::::


':: start MODULE ADMIN MENU :::::::::::::::::::::::::::::::::::::::::::::
sub article_admin_button()
ct = 3
  mINAME = "articles_admin"
  mTitle = "* Articles ADMIN"
  mName = "Articles"
  msName = "Article"
  mLink1 = "admin_articles.asp"
  mLink2 = "admin_article_main.asp"
  
  response.Write("<hr><h4>" & mName & " SAdmin Menu</h4><br>")
 
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
%>