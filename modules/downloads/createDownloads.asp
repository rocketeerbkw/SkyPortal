<!--#include file="config.asp" --><%
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
dim do_app, app_version, app_id, dirPath, tDirPath, tDirPath2, sNames
bUninstall = false
bReinstall = false

':: leave this as is.
strModTablePrefix = ""
app_version = "1.1"
do_app = true
incDlFp = false
dirPath = server.MapPath("files/downloads")
tDirPath = Server.MapPath("files") & "\tempdl"
tDirPath2 = Server.MapPath("files") & "\tempdl2"
sNames = ""
sTxt = ""
bIsUpgrade = false
%>
<!--#include file="lang/en/core_install_data.asp" -->
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<!--#INCLUDE file="includes/inc_DBFunctions.asp" -->
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
  if incDlFp then
    if request.QueryString("typ") = "success" then
	  showInstalled()
	else
	  createModule()
	end if
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the fp_dl.asp ""include"" file<br />")
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
<%

sub dl_Upgrades()
end sub
 
sub createModule()
  spThemeBlock1_open(intSkin)
  response.Write("<hr><h3>DOWNLOADS MODULE</h3><br /><br />")
    set oFs = new clsSFSO
	  sPath = server.MapPath("files/") & "/downloads"
      if bUninstall then
	    oFs.DeleteFolder(sPath)
      end if
	  if not bUninstall and not bReinstall then
	    oFs.CreateFolder(sPath)
	  end if
    set oFs = nothing
	
  if bUninstall then
	uninstall_DL()
	if ErrorCount = 0 then
	  Call setSession("Downloads","usuccess")
	else
	  sTxt = "Downloads module not Uninstalled.<br />The Module had errors during Uninstall"
	end if
  else
	if bReinstall then
	  uninstall_DL()
	  ErrorCount = 0
  	  createDownloads()
	  if ErrorCount = 0 then
	    Call setSession("Downloads","rsuccess")
	  else
	  	sTxt = "Downloads module not Reinstalled.<br />The Module had errors during install"
	    uninstall_DL()
	  end if
	else
  	  createDownloads()
	  if ErrorCount = 0 then
	    Call setSession("Downloads","success")
	  else
	    'uninstall_DL()
	    sTxt = sTxt & "Downloads module not installed.<br />The Module had errors during install"
	  end if
	end if
  end if
  mnu.DelMenuFiles("")
  
	
	if ErrorCount = 0 then
      resetCoreConfig()
	  closeAndGo(sScript & "?typ=success")
	else
	  Call setSession("Downloads","")
  	  response.Write("<hr><h4>" & sTxt & "</h4><br><br><br><br>")
	end if
  spThemeBlock1_close(intSkin)
end sub
 
sub createDownloads()
  
  'check if app is existing
  sSql = "SELECT APP_NAME, APP_ID, APP_VERSION FROM PORTAL_APPS WHERE APP_INAME = 'downloads'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
  end if
  set rsA = nothing
  
  if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "1.1"
	   ':: current version
	   ':: rel: 20080420
	    Call setSession("Downloads","upcurrent")
	    closeAndGo(sScript & "?typ=success")
     case "1.0"
	   ':: current release version
	   ':: rel: 20080420
	   delSAdminMenu()
	   updateVersion app_version,"downloads"
     case "0.12"
	   ':: rel: 20080226
	   delSAdminMenu()
	   updateVersion app_version,"downloads"
     case "0.11"
	   ':: rel: 20080215
	   dl_Upgrade11_12()
	   updateVersion app_version,"downloads"
     case "0.10"
  	   updateFileStructure()
	   dl_Upgrade10_11()
	   dl_Upgrade11_12()
	   updateVersion app_version,"downloads"
     case "0.9"
  	   updateFileStructure()
	   dl_Upgrade10_11()
	   dl_Upgrade11_12()
	   updateVersion app_version,"downloads"
	 case else
  	   'updateFileStructure()
	   'dl_Upgrades()
	   'updateVersion app_version,"downloads"
   end select
	  if ErrorCount = 0 then
	    Call setSession("Downloads","upsuccess")
		mnu.DelMenuFiles("")
        resetCoreConfig()
	    closeAndGo(sScript & "?typ=success")
	  else
	  	sTxt = "Downloads module not Upgraded.<br />The Module had errors during upgrade<br />"
	    'uninstall_DL()
	  end if
  elseif not bUninstall then
    addApp()
	dl_main_button()
    newAdminMenuDL()
    addFp()
    addUploads()
    crMainTbl()
	crStructure()
    addDMODs()
    addIntro()
    addSkypage()

	'dl_Upgrades()
  end if
end sub
	
sub addApp()
  'create the app
  response.Write("<hr><b>Update PORTAL_APPS</b><br /><br />")
  redim arrData(2)
  arrData(0) = "[PORTAL_APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_GROUPS_WRITE],[APP_GROUPS_FULL],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_CONFIG],[APP_VIEW],[APP_VERSION],[APP_SUBSEC]"
  arrData(2) = "'downloads','downloads',1,0,'1,2,3','1,2','1',1,1,'config_downloads','dl.asp','" & app_version & "',0"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM PORTAL_APPS WHERE APP_INAME = 'downloads'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub addFp()
	'add downloads to front page items
  response.Write("<hr><b>Add data to PORTAL_FP table</b><br /><br />")
	redim arrData(11)
	arrData(0) = "[PORTAL_FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Downloads - Popular','dl_popular_sm','dl_small:top',1,4,'Most popular downloads.','3'," & app_id & ""
	arrData(3) = "'Downloads - Popular','dl_popular_lg','dl_large:top',1,2,'Most popular downloads.','3'," & app_id & ""
	arrData(4) = "'Downloads - Newest','dl_newest_sm','dl_small:new',1,4,'Newest downloads.','3'," & app_id & ""
	arrData(5) = "'Downloads - Newest','dl_newest_lg','dl_large:new',1,2,'Newest downloads.','3'," & app_id & ""
	arrData(6) = "'Downloads - Random','dl_random_sm','dl_small:random',1,4,'Random downloads.','3'," & app_id & ""
	arrData(7) = "'Downloads - Random','dl_random_lg','dl_large:random',1,2,'Random downloads.','3'," & app_id & ""
	arrData(8) = "'Downloads - Featured','dl_featured_sm','dl_small:featured',1,4,'Featured downloads.','3'," & app_id & ""
	arrData(9) = "'Downloads - Featured','dl_featured_lg','dl_large:featured',1,2,'Featured downloads.','3'," & app_id & ""
	arrData(10) = "'Downloads Menu','dl_menu','menu_dl',1,4,'Default Downloads Manager menu.','1,2,3'," & app_id & ""
	arrData(11) = "'Downloads Intro','dl_intro','mod_displayIntro:" & app_id & "',1,4,'Default Downloads Intro.','1,2,3'," & app_id & ""
	populateB(arrData)
end sub

sub addUploads()

  response.Write("<hr><b>Add data to PORTAL_UPLOAD_CONFIG table</b><br /><br />")
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
		strSql = strSql & ", 'zip,rar'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '1,2'"
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
		'response.Write(strSql)
		populateA(strSql)
end sub

sub crStructure()
 bCrItem = false
 pid = addParent()
 for xc = 1 to 2
  response.Write("<hr>Adding Downloads Category.....<br>")
  sSql = "INSERT INTO " & strTablePrefix & "M_CATEGORIES ("
  sSQL = sSql & "CAT_NAME"
  sSQL = sSql & ",CAT_SDESC"
  sSQL = sSql & ",CAT_LDESC"
  sSQL = sSql & ",PARENT_ID"
  sSQL = sSql & ",CG_READ"
  sSQL = sSql & ",CG_WRITE"
  sSQL = sSql & ",CG_FULL"
  sSQL = sSql & ",CG_INHERIT"
  sSQL = sSql & ",CG_PROPAGATE"
  sSQL = sSql & ",APP_ID"
  sSQL = sSql & ",C_ORDER"
  sSQL = sSql & ")VALUES("
  select case xc
    case 1
      sSQL = sSql & "'Graphics'"
      sSQL = sSql & ",'Graphics Category'"
      sSQL = sSql & ",'Graphics Category'"
	  ssSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_NAME = 'Graphics' AND APP_ID=" & app_id
	case 2
      sSQL = sSql & "'Webpage Development'"
      sSQL = sSql & ",'Webpage Development Category'"
      sSQL = sSql & ",'Webpage Development Category'"
	  ssSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_NAME = 'Webpage Development' AND APP_ID=" & app_id
	case else
  end select
  sSQL = sSql & "," & pid & ""
  sSQL = sSql & ",'1,2,3'"
  sSQL = sSql & ",'1,2'"
  sSQL = sSql & ",'1'"
  sSQL = sSql & ",1"
  sSQL = sSql & ",1"
  sSQL = sSql & "," & app_id & ""
  sSQL = sSql & "," & xc
  sSQL = sSql & ")"
  populateA(sSQL)
	
	set rsA = my_Conn.execute(ssSql)
	c_id = rsA(0)
	c_news_cat = c_id
	set rsA = nothing
	
  for xm = 1 to 4
	bCrItem = false
    select case xc
      case 1
    	select case xm
	  	  case 1
	    	sxName = "Screensavers"
	  	  case 2
	    	sxName = "Themes"
	  	  case 3
	    	sxName = "Wallpapers"
	  	  case 4
	    	sxName = "Web Graphics"
		end select
      case 2
    	select case xm
	  	  case 1
	    	sxName = "ASP"
			bCrItem = true
	  	  case 2
	    	sxName = "PHP"
	  	  case 3
	    	sxName = "JavaScript"
	  	  case 4
	    	sxName = "Others"
		end select
	  case else
	    sxName = "ERROR"
	end select
    response.Write("<br><br>Adding SubCategory.....<br>")
	':: create news subcats
    sSql = "INSERT INTO " & strTablePrefix & "M_SUBCATEGORIES ("
    sSQL = sSql & "SUBCAT_NAME"
    sSQL = sSql & ",SUBCAT_SDESC"
    sSQL = sSql & ",SUBCAT_LDESC"
    sSQL = sSql & ",CAT_ID"
    sSQL = sSql & ",SG_READ"
    sSQL = sSql & ",SG_WRITE"
    sSQL = sSql & ",SG_FULL"
    sSQL = sSql & ",SG_INHERIT"
    'sSQL = sSql & ",SUBCAT_IMAGE"
    sSQL = sSql & ",APP_ID"
    sSQL = sSql & ",C_ORDER"
	if bCrItem then
      sSQL = sSql & ",ITEM_CNT"
	end if
    sSQL = sSql & ")VALUES("
  
    tSQL = sSQL & "'" & sxName & "'"
    tSQL = tSQL & ",'" & sxName & " SubCategory'"
    tSQL = tSQL & ",'" & sxName & " SubCategory'"
    tSQL = tSQL & "," & c_id & ""
    tSQL = tSQL & ",'1,2,3'"
    tSQL = tSQL & ",'1,2'"
    tSQL = tSQL & ",'1'"
    tSQL = tSQL & ",1"
    'tSQL = tSQL & ",''"
    tSQL = tSQL & "," & app_id & ""
    tSQL = tSQL & "," & xm
	if bCrItem then
      tSQL = tSQL & ",1"
	end if
    tSQL = tSQL & ")"
    populateA(tSQL)
	
	if bCrItem then
      response.Write("<br><br>Creating Subcategory Item.....<br>")
	  sSql = "SELECT SUBCAT_ID FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_NAME = '" & sxName & "' AND APP_ID=" & app_id
	  set rsA = my_Conn.execute(sSql)
	  sc_id = rsA(0)
	  set rsA = nothing
	
      call crItem(sc_id,xc)
	end if
  next
 next 
end sub
  
function addParent()
  response.Write("<hr><b>Populate Parent table</b><br /><br />")
  ':: insert parent table row
  t_id = 0
  sSql = "INSERT INTO " & strTablePrefix & "M_PARENT ("
  sSQL = sSql & "PARENT_NAME,PARENT_SDESC,PARENT_LDESC"
  sSQL = sSql & ",PG_READ,PG_WRITE,PG_FULL,PG_INHERIT"
  sSQL = sSql & ",PG_PROPAGATE,APP_ID"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'Default Downloads'"
  sSQL = sSql & ",'Default Downloads parent group'"
  sSQL = sSql & ",'Default Downloads parent group'"
  sSQL = sSql & ",'1,2,3','1,2','1',1,1," & app_id & ""
  sSQL = sSql & ")"
  executeThis(sSQL)
  
  sSql = "SELECT PARENT_ID FROM " & strTablePrefix & "M_PARENT WHERE PARENT_NAME = 'Default Downloads'"
  set rsA = my_Conn.execute(sSql)
  t_id = rsA(0)
  set rsA = nothing
  addParent = t_id
end function

sub crMainTbl()
'::::::::::::::::::: CREATE DL TABLE ::::::::::::::::::
  response.Write("<hr><b>Create DOWNLOAD table</b><br /><br />")
sSQL = "CREATE TABLE [DL]([ACTIVE] LONG, [BADLINK] LONG, [CATEGORY] LONG, [CONTENT] MEMO, [DESCRIPTION] TEXT(255), [DL_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [EMAIL] TEXT(100), [FEATURED] INT NOT NULL DEFAULT 0, [FILESIZE] TEXT(100), [HIT] LONG DEFAULT 0, [KEYWORD] TEXT(255), [LANG] TEXT(100), [NAME] TEXT(100), [O_POST_DATE] TEXT(25), [POST_DATE] TEXT(100), [RATING] LONG NOT NULL DEFAULT 0, [TDATA1] TEXT(255), [TDATA10] TEXT(255), [TDATA2] TEXT(255), [TDATA3] TEXT(255), [TDATA4] TEXT(255), [TDATA5] TEXT(255), [TDATA6] TEXT(255), [TDATA7] TEXT(255), [TDATA8] TEXT(255), [TDATA9] TEXT(255), [UPDATED] TEXT(50) DEFAULT 0, [UPLOADER] TEXT(100), [URL] TEXT(200), [VOTES] LONG DEFAULT 0);"

createTable(checkIt(sSQL))
end sub

sub crItem(s,n)
	'-------------------- populate table with default values --------------------------
	strSql = "INSERT INTO DL"
	strSql = strSql & "(NAME"
	strSql = strSql & ", URL"
	strSql = strSql & ", KEYWORD"
	strSql = strSql & ", CATEGORY"
	strSql = strSql & ", DESCRIPTION"
	strSql = strSql & ", CONTENT"
	strSql = strSql & ", EMAIL"
	strSql = strSql & ", POST_DATE"
	strSql = strSql & ", O_POST_DATE"
	strSql = strSql & ", ACTIVE"
	strSql = strSql & ", BADLINK"
	strSql = strSql & ", FEATURED"
	strSql = strSql & ", FILESIZE"
	strSql = strSql & ", UPLOADER"
	strSql = strSql & ", TDATA1"
	strSql = strSql & ", TDATA2"
	strSql = strSql & ", TDATA3"
	strSql = strSql & ", TDATA4"
	strSql = strSql & ", TDATA5"
	strSql = strSql & ")"
	strSql = strSql & " VALUES ("
	strSql = strSql & "'SkyPortal v" &  strWebSiteMVersion & "'"
	strSql = strSql & ", 'http://www.SkyPortal.net'"
	strSql = strSql & ", 'SkyPortal,portal'"
	strSql = strSql & ", '" & s & "'"
	strSql = strSql & ", 'SkyPortal v" &  strWebSiteMVersion & ": Your complete website portal solution'"
	strSql = strSql & ", 'SkyPortal v" &  strWebSiteMVersion & ": Your complete website portal solution'"
	strSql = strSql & ", 'nospam@nospam.net'"
	strSql = strSql & ", '" & strCurDateString & "'"
	strSql = strSql & ", '" & strCurDateString & "'"
	strSql = strSql & ", " & "1"
	strSql = strSql & ", " & "0"
	strSql = strSql & ", " & "1"
	strSql = strSql & ", " & "'2mb'"
	strSql = strSql & ", " & "'" & strDBNTUserName & "'"
	strSql = strSql & ", " & "'SkyPortal EULA'"
	strSql = strSql & ", " & "'Multi Language'"
	strSql = strSql & ", " & "'Windows'"
	strSql = strSql & ", " & "'SkyPortal.net'"
	strSql = strSql & ", " & "'http://www.SkyPortal.net'"
	strSql = strSql & ")"

	'executeThis(strSQL)
	populateA(strSql)
end sub

sub uninstall_DL()
  response.Write("<hr><b>Uninstall App</b><br /><br />")
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'downloads'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
	apid = rsA("APP_ID")
  else
    exit sub
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
  
  droptable("DL")
	
  sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_INAME='downloads'"
  executeThis(sSql)
  response.Write("<b>Module Uninstall Complete</b><br /><hr><br />")
end sub

sub migrateIntegratedTables()
  response.Write("<hr><hr><b>Migrate Integrated Module Tables</b>")
  response.Write("<br/><br/>")
  response.Write("<b>Migrate Integrated Categories and Subcats</b>")
  response.Write("<br /><br />")
  
  ':: insert parent table row
  sSql = "INSERT INTO " & strTablePrefix & "M_PARENT ("
  sSQL = sSql & "PARENT_NAME,PARENT_SDESC,PARENT_LDESC"
  sSQL = sSql & ",PG_READ,PG_WRITE,PG_FULL,PG_INHERIT"
  sSQL = sSql & ",PG_PROPAGATE,APP_ID"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'Default Downloads'"
  sSQL = sSql & ",'Default Downloads parent group'"
  sSQL = sSql & ",'Default Downloads parent group'"
  sSQL = sSql & ",'1,2,3','1,2','1',1,1,0"
  sSQL = sSql & ")"
  executeThis(sSQL)
  
  strSql = "ALTER TABLE DL ADD [INTEGRATED] INT;"
  alterTable(checkIt(strSql))
  sSql = "UPDATE DL SET INTEGRATED = 0"
  executeThis(sSql)
  
  sSql = "SELECT * FROM DL_CATEGORIES ORDER BY CAT_ID"
  set rsC = my_Conn.execute(sSql)
  if not rsC.eof then
    do until rsC.eof
	  oldCatID = rsC("CAT_ID")
	  newCatID = integrateMCategory(rsC)
	  
  	  sSql = "SELECT * FROM DL_SUBCATEGORIES WHERE CAT_ID = " & oldCatID & " ORDER BY SUBCAT_ID"
  	  set rsS = my_Conn.execute(sSql)
  	  if not rsS.eof then
    	do until rsS.eof
	  	  oldSCatID = rsS("SUBCAT_ID")
	      newSCatID = integrateMSubCategory(rsS,newCatID)
 		  if bFSO then
		    scTo = dirPath & "\" & newSCatID
			scFrm = tDirPath2 & "\" & oldSCatID
			moveFolder scFrm,scTo
		  end if
		  
		  sSql = "SELECT * FROM DL WHERE CATEGORY=" & oldSCatID
		  set rsF = my_Conn.execute(sSql)
		  if not rsF.eof then
		    do until rsF.eof
	  		  sURL = rsF("URL")
	  		  if instr(sURL,strHomeURL) > 0 then
	    	    sURL = replace(sURL,strHomeURL,"")
	  		  end if
			  ' files/downloads/23/2007.zip
			  ' files/downloads/subcat/item.zip
	  		  if left(sURL,16) = "files/downloads/" then
				if instr(sURL,"downloads/" & oldSCatID & "/") > 0 then
				  sURL = replace(sURL,"downloads/" & oldSCatID & "/","downloads/" & newSCatID & "/")
				end if
	  
  	  	  		sSql = "UPDATE DL SET CATEGORY="& newSCatID
  	  	  		sSql = sSql & ", INTEGRATED = 1"
  	  	  		sSql = sSql & ", URL = '" & sURL & "'"
  	  	  		sSql = sSql & " WHERE DL_ID = " & rsF("DL_ID")
  	  	  		sSql = sSql & " AND INTEGRATED = 0"
				executeThis(sSql)
			  else
  	  	  		sSql = "UPDATE DL SET CATEGORY="& newSCatID
  	  	  		sSql = sSql & ", INTEGRATED = 1"
  	  	  		sSql = sSql & " WHERE DL_ID = " & rsF("DL_ID")
  	  	  		sSql = sSql & " AND INTEGRATED = 0"
				executeThis(sSql)
	  		  end if
			  rsF.movenext
			loop
		  end if
		  set rsF = nothing
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
  sSql = "SELECT * FROM DL_RATING ORDER BY RATING_ID"
  set rsR = my_Conn.execute(sSql)
  if not rsR.eof then
    do until rsR.eof
  	  integrateMRatings(rsR)
	  rsR.movenext
	loop
  end if
  set rsR = nothing
  
  ':: clean up
  strSql = "ALTER TABLE DL DROP COLUMN [INTEGRATED];" 
  alterTable(checkIt(strSql))
  
  droptable("DL_CATEGORIES")
  droptable("DL_SUBCATEGORIES")
  droptable("DL_RATING")
  
  deleteFolder(tDirPath2)
end sub

sub integrateMRatings(obj)
  sSql = "INSERT INTO " & strTablePrefix & "M_RATING ("
  sSQL = sSql & "ITEM_ID,RATE_BY,RATE_DATE"
  sSQL = sSql & ",RATING,COMMENTS,APP_ID"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & obj("DL")
  sSQL = sSql & "," & obj("RATE_BY")
  sSQL = sSql & ",'" & obj("RATE_DATE") & "'"
  sSQL = sSql & "," & obj("RATING")
  sSQL = sSql & ",'" & replace(obj("COMMENTS"),"'","") & "'"
  sSQL = sSql & "," & app_id & ""
  sSQL = sSql & ")"
  executeThis(sSQL)
end sub

function integrateMSubCategory(obj,cat)
  tCount = getCount("DL_ID","DL","CATEGORY=" & obj("SUBCAT_ID") & "")
  
  response.Write("<b>Migrate Integrated Subcategories</b><br /><br />")
  sSql = "INSERT INTO " & strTablePrefix & "M_SUBCATEGORIES ("
  sSQL = sSql & "SUBCAT_NAME,SUBCAT_SDESC,SUBCAT_LDESC"
  sSQL = sSql & ",CAT_ID,SG_READ,SG_WRITE,SG_FULL"
  sSQL = sSql & ",SG_INHERIT,APP_ID,C_ORDER,ITEM_CNT"
  sSQL = sSql & ")VALUES("
  sSQL = sSql & "'" & obj("SUBCAT_NAME") & "'"
  sSQL = sSql & ",'" & obj("SUBCAT_DESC") & "'"
  sSQL = sSql & ",'" & obj("SUBCAT_DESC") & "'"
  sSQL = sSql & "," & cat
  sSQL = sSql & ",'" & obj("SG_READ") & "'"
  sSQL = sSql & ",'" & obj("SG_WRITE") & "'"
  sSQL = sSql & ",'" & obj("SG_FULL") & "'"
  sSQL = sSql & "," & obj("SG_INHERIT")
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
  sSQL = sSql & ",'" & obj("CAT_DESC") & "'"
  sSQL = sSql & ",'" & obj("CAT_DESC") & "'"
  sSQL = sSql & ",'" & obj("CG_READ") & "'"
  sSQL = sSql & ",'" & obj("CG_WRITE") & "'"
  sSQL = sSql & ",'" & obj("CG_FULL") & "'"
  sSQL = sSql & "," & obj("CG_INHERIT") & ""
  sSQL = sSql & "," & obj("CG_PROPAGATE") & ""
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

sub updateFileStructure()
 if bFSO then
  on error resume next
  set fso = Server.CreateObject("Scripting.FileSystemObject")
  'dirPath = server.MapPath("files/downloads")
  'tDirPath = Server.MapPath("files") & "\tempdl"
  'tDirPath2 = Server.MapPath("files") & "\tempdl2"
  'sNames = ""
	
  ':: move to temp directory
  fso.createfolder(tDirPath)
  fso.createfolder(tDirPath2)
  set oF = fso.getfolder(dirPath)
  for each f in oF.SubFolders
	set fo=fso.GetFolder(f.path)
    fo.move(tDirPath & "\" & f.name)
	set fo = nothing	  
  next
  set oF = nothing
  
  ':: move subcats back to tDirPath2 directory
  set oF = fso.getfolder(tDirPath)
  for each f in oF.SubFolders
    sNames = sNames & f.name & ","
	set cf = fso.GetFolder(f.path)
	for each sf in cf.SubFolders
	  set tf = fso.GetFolder(sf.path)
	  tf.Move(tDirPath2 & "\" & sf.name)
	  set tf = nothing
	next
	set cf = nothing	  
  next
  set f = nothing
  set oF = nothing
  set fso = nothing
  on error goto 0
  
  ':: clean up - delete temp directory
  deleteFolder(tDirPath)
  
  ':: do the database work
  ':: delete the categories from the DL URL fields
 if sNames <> "" then
  sNames = left(sNames,len(sNames)-1)
  response.Write(sNames)
  arNames = split(sNames,",")
  sSql = "SELECT URL, DL_ID FROM DL"
  set rsA = my_Conn.execute(sSql)
  if not rsA.eof then
    do until rsA.eof
	  sURL = rsA("URL")
	  if instr(sURL,strHomeURL) > 0 then
	    sURL = replace(sURL,strHomeURL,"")
	  end if
	  if left(sURL,16) = "files/downloads/" then
	    for x = 0 to ubound(arNames)
	      if instr(sURL,"/downloads/" & arNames(x) & "/") > 0 then
	        tUrl = replace(sURL,"/downloads/" & arNames(x) & "/","/downloads/")
	        sSql = "UPDATE DL SET URL='" & tUrl & "'"
			sSql = sSql & " WHERE DL_ID = " & rsA("DL_ID")
			executeThis(sSql)
		  end if
		next
	  end if
	  rsA.movenext
	loop
  end if
  set rsA = nothing
 end if
 end if
end sub

sub dl_Upgrade11_12()
  response.Write("<hr><hr><b>UPDATE DOWNLOADS MODULE From v0.11 to v0.12</b><br />")
  response.Write("<hr><b>Update DL table</b><br /><br />")
  
  sSql = "ALTER TABLE DL ADD [O_POST_DATE] TEXT(25) NULL"
  alterTable(checkIt(sSql))
  
  sSql = "UPDATE DL SET DL.O_POST_DATE = DL.POST_DATE"
  executeThis(sSql)
  
  sSql = "UPDATE DL SET DL.O_POST_DATE = DL.UPDATED WHERE DL.UPDATED <> '0'"
  executeThis(sSql)
  
  sSql = "SELECT SUBCAT_ID FROM PORTAL_M_SUBCATEGORIES"
  sSql = sSql & " WHERE APP_ID=" & app_id
  set rsA = my_Conn.execute(sSql)
  if not rsA.eof then
    do until rsA.eof
	  iSid = rsA("SUBCAT_ID")
      rsA.movenext
	  iCnt = getCount("POST_DATE","DL","CATEGORY=" & iSid & "")
	  sSql = "UPDATE PORTAL_M_SUBCATEGORIES SET ITEM_CNT=" & iCnt
	  sSql = sSql & " WHERE SUBCAT_ID=" & iSid
	  executeThis(sSql)
	  'Response.Write sSql & "<br><br>"
    loop
  end if
  set rsA = nothing
  
  delSAdminMenu()
end sub

sub dl_Upgrade10_11()
  response.Write("<hr><hr><b>UPDATE DOWNLOADS MODULE From v0.10 to v0.11</b><br />")
  response.Write("<hr><b>Update " & strTablePrefix & "APPS table</b><br /><br />")
  strSql = "UPDATE "&strTablePrefix&"APPS SET APP_GROUPS_USERS = '1,2,3'"
  strSql = strSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
  strSql = strSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & DateToStr(now()) & "'"
  strSql = strSql & ", APP_SUBSEC = 0 WHERE APP_INAME = 'downloads';"
  executeThis(strSql)
  
  response.Write("<hr><b>Update DL_CATEGORIES table</b><br /><br />")
  strSql = "UPDATE DL_CATEGORIES SET CG_READ = '1,2,3'"
  strSql = strSql & ",CG_WRITE = '1,2', CG_FULL = '1'"
  strSql = strSql & ",CG_INHERIT = 1, CG_PROPAGATE = 1;"
  executeThis(strSql)
  
  strSql = "ALTER TABLE DL_CATEGORIES ADD [CAT_IMAGE] TEXT(255) NULL, [CAT_DESC] MEMO NULL;"
  alterTable(checkIt(strSql))
  
  response.Write("<hr><b>Update DL_SUBCATEGORIES table</b><br /><br />")
  strSql = "UPDATE DL_SUBCATEGORIES SET SG_READ = '1,2,3'"
  strSql = strSql & ",SG_WRITE = '1,2', SG_FULL = '1'"
  strSql = strSql & ",SG_INHERIT = 1;"
  executeThis(strSql)
  
  strSql = "ALTER TABLE DL_SUBCATEGORIES ADD [SUBCAT_IMAGE] TEXT(255) NULL, [SUBCAT_DESC] MEMO NULL;"
  alterTable(checkIt(strSql))
  
  Response.Write("<br /><b>Alter DL Fields</b><br />")
  strSql = "ALTER TABLE DL ADD [ACTIVE] INT"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [CONTENT] MEMO NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [UPDATED];"
  doSQL2 checkIt(strSql),1
  strSql = "ALTER TABLE DL ADD [UPDATED] TEXT(50) DEFAULT 0"
  alterTable(checkIt(strSql))
  sSql = "ALTER TABLE DL ALTER COLUMN KEYWORD TEXT(255) NULL"
  doSQL2 checkIt(sSql),1
	
  addTData()
  
  Response.Write("<br /><b>Update DL CONTENT</b><br />")
  strSql = "UPDATE DL SET ACTIVE = 1 WHERE SHOW = 1;"
  executeThis(strSql)
  strSql = "UPDATE DL SET ACTIVE = 0 WHERE SHOW = 0;"
  executeThis(strSql)
  strSql = "UPDATE DL SET UPDATED = 0;"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.CONTENT = DL.DESCRIPTION WHERE DL_ID > 0"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.TDATA1 = DL.LICENSE"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.TDATA2 = DL.LANG"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.TDATA3 = DL.PLATFORM"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.TDATA4 = DL.PUBLISHER"
  executeThis(strSql)
  strSql = "UPDATE DL SET DL.TDATA5 = DL.PUBLISHER_URL"
  executeThis(strSql)
  
  Response.Write("<br /><b>Drop column</b><br />")
  strSql = "ALTER TABLE DL DROP COLUMN [SHOW];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [LICENSE];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [LANG];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [PLATFORM];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [PUBLISHER];" 
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL DROP COLUMN [PUBLISHER_URL];" 
  alterTable(checkIt(strSql))
  
  addDMODs()
  addIntro()
  addNewFP()
  addSkypage()
  newAdminMenuDL()
  migrateIntegratedTables()
end sub

sub addDMODs()
    Response.Write("<br /><b>Add Pending Tasks & Site Search</b><br />")
	redim arrData(4)
	arrData(0) = strTablePrefix & "MODS"
	arrData(1) = "M_NAME, M_CODE, M_VALUE"
	arrData(2) = "'" & app_id & "', 'admTaskLnk', 'dl_adminPndLink()'"
	arrData(3) = "'" & app_id & "', 'siteSrch', 'dl_SiteSearch()'"
	arrData(4) = "'" & app_id & "', 'pndTskCnt', 'dl_PendTaskCnt()'"
	populateB(arrData)
end sub

sub addTData()
  strSql = "ALTER TABLE DL ADD [TDATA1] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA2] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA3] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA4] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA5] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA6] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA7] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA8] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA9] TEXT(255) NULL"
  alterTable(checkIt(strSql))
  strSql = "ALTER TABLE DL ADD [TDATA10] TEXT(255) NULL"
  alterTable(checkIt(strSql))
end sub

sub modifyNulls()
end sub

sub addNewFP()
  response.Write("<br><b>Add data to PORTAL_FP table</b><br />")
	redim arrData(3)
	arrData(0) = "[PORTAL_FP]"
	arrData(1) = "[FP_NAME],[FP_INAME],[FP_FUNCTION],[FP_ACTIVE],[FP_COLUMN],[FP_DESC],[FP_GROUPS],[APP_ID]"
	arrData(2) = "'Downloads Menu','dl_menu','menu_dl',1,4,'Default Downloads Manager menu.','1,2,3'," & app_id & ""
	arrData(3) = "'Downloads Intro','dl_intro','mod_displayIntro:" & app_id & "',1,4,'Default Downloads Intro.','1,2,3'," & app_id & ""
	populateB(arrData)
end sub

sub addSkypage()
	
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
		strSql = strSql & "'Download Manager'"  'P_NAME
		strSql = strSql & ", 'downloads'" 	'P_INAME
		strSql = strSql & ", 'Download Manager'" 'P_TITLE
		strSql = strSql & ", ' '" 'P_CONTENT
		strSql = strSql & ", ' '" 'P_ACONTENT
		strSql = strSql & ", 'Downloads Menu:menu_dl,Downloads - Newest:dl_small:new'" 'P_LEFTCOL
		strSql = strSql & ", 'Downloads - Featured:dl_small:featured,Downloads - Popular:dl_small:top,Downloads - Random:dl_small:random'" 'P_RIGHTCOL
		strSql = strSql & ", 'Downloads Intro:mod_displayIntro:" & app_id & "'" 'P_MAINTOP
		strSql = strSql & ", 'Downloads - Newest:dl_large:new'" 'P_MAINBOTTOM
		strSql = strSql & ", " & app_id & ""  'P_APP
		strSql = strSql & ", 0"  'P_USE_PG_DISP
		strSql = strSql & ", 'dl.asp'" 'P_OTHER_URL
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
  sImsg = "<b>Welcome to the <span class=""fAlert""><b>NEW</b></span> SkyPortal <i>Downloads Manager</i> Module.</b><br/>"
  sImsg = sImsg & "<br/>This is your module introduction block. You can create any message that you want. Your visitors will only see this message on the module main page. You can edit this message by clicking the *edit* icon in the title bar above this message."
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
		strSql = strSql & "'Downloads Intro'"
		strSql = strSql & ", 'Downloads Introduction'"
		strSql = strSql & ", '" & sImsg & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", " & app_id
		strSql = strSql & ", 1"						
		strSql = strSql & ")"
'		response.Write(strSql)
		populateA(strSql)
end sub

sub delSAdminMenu()
  sSql = "DELETE FROM menu where INAME = 'downloads_admin'"
  'executeThis(sSql)
  doSQL2 checkIt(sSql),1
end sub

sub newAdminMenuDL()
  delSAdminMenu()
  
  mINAME = "downloads_admin"
  mTitle = "* Downloads ADMIN"
  mName = "Downloads"
  msName = "Download"
  mLink1 = "admin_dl_admin.asp"
  mLink2 = "admin_dl_main.asp"
  
  response.Write("<hr><h4>" & mName & " SAdmin Menu</h4><br />")
 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuOrder"
arrData(2) = "'" & mName & "', '" & mINAME & "','','','','','" & mTitle & "','" & mINAME & "',"& app_id &",1"
'populateB(arrData)

'sSql = "select ID from menu where Name = '"& mName &"' and INAME = '"& mINAME &"'"
'set rsT = my_Conn.execute(sSql)
'pID = rsT(0)
'set rsT = nothing

redim arrData(4)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,Target,mnuImage,onClick,mnuTitle,INAME,ParentID,app_id,mnuOrder"

arrData(2) = "'Attention Items', '" & mName & "','" & mLink1 & "','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",1"
arrData(3) = "'Category Manager', '" & mName & "','" & mLink1 & "?cmd=20','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",2"
arrData(4) = "'Subcategory Manager', '" & mName & "','" & mLink1 & "?cmd=21','_parent','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",3"

'populateB(arrData)

':: ADD to Module Admin menu 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'm_admin','','','','','Module Admin','m_admin',"& app_id &",'" & mINAME & "',1"
'populateB(arrData)
  
end sub

sub dl_Upgrade09()
  response.Write("<hr><hr><b>UPDATE DOWNLOADS MODULE</b><br />")
  response.Write("<hr><b>Update " & strTablePrefix & "APPS table</b><br /><br />")
	   strSql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_USERS = '1,2,3'"
	   strSql = strSql & ",APP_GROUPS_WRITE = '1,2', APP_GROUPS_FULL = '1'"
	   strSql = strSql & ",APP_VERSION = '" & app_version & "', APP_DATE = '" & DateToStr(now()) & "'"
	   strSql = strSql & ", APP_SUBSEC = 0 WHERE APP_INAME = 'downloads';"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update DL_CATEGORIES table</b><br /><br />")
	   strSql = "DL_CATEGORIES"
	   strSql = strSql & ",[CG_READ] MEMO NULL,[CG_WRITE] MEMO NULL,[CG_FULL] MEMO NULL,[CG_INHERIT] INT NULL DEFAULT 1,[CG_PROPAGATE] INT NULL DEFAULT 1"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE DL_CATEGORIES "
	   strSql = strSql & "SET CG_READ = '1,2,3', CG_WRITE = '1,2', CG_FULL = '1', CG_INHERIT = 1 "
	   strSql = strSql & "WHERE CAT_ID not like 0;"
	   executeThis(strSql)
	   
  response.Write("<hr><b>Update DL_SUBCATEGORIES table</b><br /><br />")
	   strSql = "DL_SUBCATEGORIES"
	   strSql = strSql & ",[SG_READ] MEMO NULL,[SG_WRITE] MEMO NULL,[SG_FULL] MEMO NULL,[SG_INHERIT] INT NULL"
	   alterTable2(checkIt(strSql))
	   
	   strSql = "UPDATE DL_SUBCATEGORIES "
	   strSql = strSql & "SET SG_READ = '1,2,3', SG_WRITE = '1,2', SG_FULL = '1', SG_INHERIT = 1 "
	   strSql = strSql & "WHERE SUBCAT_ID not like 0;"
	   executeThis(strSql)
	
  response.Write("<hr><b>Update " & strTablePrefix & "FP table</b><br /><br />")
	   strSql= "UPDATE " & strTablePrefix & "FP SET fp_groups = '1,2,3' WHERE APP_ID = " & app_id & ";"
	   executeThis(strSql)
	
	'dl_main_button()
	'dl_admin_button()
end sub


':: start MODULE MENUS :::::::::::::::::::::::::::::::::::::::::::::::::
sub dl_main_button()
ct = 3
  mTitle = "Downloads"		' Friendly menu name
  mINAME = "m_downloads"	' Internal menu name = app_INAME : Must me different from mName
  mName = "Downloads"		' Link Head Text
  msName = "Download"
  mCntFunct = "cntNewDL()"
  mLink1 = "dl.asp"		'Main Directory
  mLink2 = "dl.asp?cmd=3"	'New
  mLink3 = "dl.asp?cmd=4"	'Popular
  mLink4 = "dl.asp?cmd=5"	'Top
  mLink5 = "dl_add_form.asp"	'Submit
  mLink6 = "openWindow3(''dl_pop.asp?mode=12'')"	'FAQ
  response.Write("<hr><h4>" & mName & " Menu</h4><br>")
 
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
arrData(4) = "'Popular "& mName &"', '" & mName & "','" & mLink3 & "','_parent','','','','" & mTitle & "','" & mINAME & "'," & pID & ","& app_id &",'1,2,3',3"
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
sub dl_admin_button()
ct = 3
  mINAME = "downloads_admin"
  mTitle = "* Downloads ADMIN"
  mName = "Downloads"
  msName = "Download"
  mLink1 = "admin_dl_admin.asp"
  mLink2 = "admin_dl_main.asp"
  
  response.Write("<hr><h4>" & mName & " SAdmin Menu</h4><br />")
 
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuOrder"
arrData(2) = "'" & mName & "', '" & mINAME & "','','','','','" & mTitle & "','" & mINAME & "',"& app_id &",1"
'populateB(arrData)

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

'populateB(arrData)
 
 ':: add module links to module admin menu
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'm_admin','','','','','Module Admin','m_admin',"& app_id &",'" & mINAME & "',1"
'populateB(arrData)
 
 ':: add module links to superadmin menu
redim arrData(2)
arrData(0) = "Menu"
arrData(1) = "Name,Parent,Link,mnuImage,onClick,Target,mnuTitle,INAME,app_id,mnuAdd,mnuOrder"
arrData(2) = "'" & mTitle & " Menu', 'sadmin','','','','','Portal Admin','sadmin',"& app_id &",'" & mINAME & "',4"
'populateB(arrData)

end sub
':: end MODULE ADMIN menu ::::::::::::::::::::::::::::::::::::::::::::::::

sub showInstalled()
  sessDownloads = readSession("Downloads")
  
  spThemeTitle = "<center>SkyPortal Downloads Module v" & app_version & "</center>"
  spThemeBlock1_open(intSkin)
  Response.Write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">"
  Response.Write "<tr><td align=""center"" colspan=""2"">"
  Response.Write "<hr>"
  Response.Write "</td></tr>"
  
  ':: skynews module install check
  select case sessDownloads
    case "success"
	  sIcn = icon(icnCheck,"","","","")
	  sTxt = "Downloads module successfully installed!"
	case "usuccess"
	  sIcn = icon(icnCheck,"","","","")
	  sTxt = "Downloads module successfully Uninstalled!"
	case "rsuccess"
	  sIcn = icon(icnCheck,"","","","")
	  sTxt = "Downloads module successfully Reinstalled!"
	case "upsuccess"
	  sIcn = icon(icnCheck,"","","","")
	  sTxt = "Downloads module successfully Upgraded to v " & app_version & "!"
	case "upcurrent"
	  sIcn = icon(icnCheck,"","","","")
	  sTxt = "Downloads module already current - v " & app_version & "!"
	case else
	  'sIcn = icon(icnDelete,"","","","")
	  'sTxt = "Downloads module not installed."
	  Call setSession("Downloads","")
	  closeAndGo(sScript)
  end select
  
  Response.Write "<tr><td align=""right"" valign=""middle"" width=""33%"">"
  Response.Write sIcn
  Response.Write "</td><td>"
  Response.Write sTxt
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""center"" colspan=""2"">"
  Response.Write "<hr>"
  Response.Write "</td></tr>"
  
  Response.Write "</table>"
  
  response.Write("<b>Be sure to delete this file '" & sScript & "' from your server!</b><br><br>")
  if sessDownloads <> "usuccess" and sessDownloads <> "" then
    response.Write("<a href=""dl.asp"">")
  else
    response.Write("<a href=""default.asp"">")
  end if
  response.Write("<b>Continue</b></a><br /><br /><br />&nbsp;")

  spThemeBlock1_close(intSkin)
  Call setSession("Downloads","")
end sub

%>
<!--#INCLUDE file="inc_footer.asp" -->
