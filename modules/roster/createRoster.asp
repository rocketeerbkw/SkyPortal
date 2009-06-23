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

'/**
' * SkyPortal Roster Module
' *
' * This file builds the database structure needed for roster module
' *   as well as adding necessary application hooks into SkyPortal
' * Or it deletes everything from the database related to the roster
' *   module... whatever floats your boat
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */


dim do_app, app_version, app_name, app_id
bUninstall = false
bReInstall = false

':: leave this as is.
strModTablePrefix = ""
app_version = "1.2"
app_name = "roster"
do_app = true
%>
<!--#INCLUDE file="config.asp" -->
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
  if incRosterFp then
    createRoster()
  else
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the fp_roster.asp ""include"" file<br>")
    Response.Write("to your fp_custom.asp file in order<br>")
    Response.Write("to install this module</p><p>&nbsp;</p>")
    spThemeBlock1_close(intSkin)
  end if
else
  spThemeBlock1_open(intSkin)
  Response.Write("<p>&nbsp;</p><p>You must be logged in as a <b>Super Admin</b>")
  Response.Write(" in order to install this module</p><p>&nbsp;</p>")
  spThemeBlock1_close(intSkin)
end if

%>
	</td>
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" --><%

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::		SUBROUTINES BELOW HERE
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub createRoster()
  spThemeBlock1_open(intSkin) %>
  <h1>ROSTER MODULE</h1>
<%  
  'check if app is existing
  sSql = "SELECT APP_NAME,APP_ID,APP_VERSION FROM " & strTablePrefix & "APPS WHERE APP_iNAME = '" & app_name & "'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall or bReInstall then
      uninstall_Roster()
	  response.Write("<h2>Roster Module Uninstalled</h2><br><br>")
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing

 if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "1.2"
      'current version
     case "1.1"
      updateRoster("1.2")
     case "1.0"
      updateRoster("1.1")
      updateRoster("1.2")
    case else
        response.write "trying to upgrade from " & cur_appVer
   end select
 elseif not bUninstall or bReInstall then
    response.Write("<h3>Install Roster</h3>")
    roster_addApp()
    roster_createTables()
	roster_skyPage()
    roster_menus()
    updateRoster("1.1")
    updateRoster("1.2")
    if request("adddata") = "1" then
    	call rosterAddDefaultData()
    end if

    Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
	response.Write("<h2>Roster Module Installed</h2><br><br>")
 end if
  response.Write("<p><b>Be sure to delete this file (createRoster.asp) from your server!</b></p>")
  response.Write("<p><a href=""default.asp""><b>Continue</b></a></p>")
  spThemeBlock1_close(intSkin)
end sub
	
sub roster_addApp()
  'create the app
  response.Write("<h3>Update PORTAL_APPS</h3>")
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_iNAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_GROUPS_WRITE],[APP_GROUPS_FULL],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_SUBSEC],[APP_CONFIG],[APP_VIEW],[APP_VERSION],[APP_DATE],[APP_tDATA1]"
  arrData(2) = "'" & app_name & "','" & app_name & "',1,0,'1,2','1,2','1',3,3,0,'config_Roster','roster.asp','" & app_version & "','" & datetostr2(now()) & "',''"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = '" & app_name & "'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
  
    'Add to upload_config
    redim arrData(2)
    arrData(0) = "[" & strTablePrefix & "UPLOAD_CONFIG]"
    arrData(1) = "[UP_ACTIVE],[UP_ALLOWEDEXT],[UP_APPID],[UP_LOCATION],[UP_LOGFILE],[UP_LOGUSERS],[UP_SIZELIMIT],[UP_THUMB_MAX_W],[UP_THUMB_MAX_H],[UP_NORM_MAX_W],[UP_NORM_MAX_H],[UP_RESIZE],[UP_CREATE_THUMB],[UP_FOLDER],[UP_ALLOWEDGROUPS]"
    arrData(2) = "1,'gif,jpg,png'," & app_id & ",'roster','upload.txt',1,5000,120,120,500,500,0,0,'files/roster/','1,2'"
    populateB(arrData)
end sub

sub roster_createTables()
	response.write("<h3>create DIVISION table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "DIVISION]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [DIVISION] TEXT(50) NOT NULL, [DESCRIP] MEMO NULL, [STARTAGE] BYTE NOT NULL, [ENDAGE] BYTE NOT NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create LEAGUE table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "LEAGUE]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [LEAGUE] TEXT(50) NOT NULL, [DESCRIP] MEMO NULL, [WEBSITE] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create PLAYER table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "PLAYER]([ID] LONG PRIMARY KEY NOT NULL, [FIRSTNAME] TEXT(50) NOT NULL, [LASTNAME] TEXT(50) NOT NULL, [SEX] TEXT(1) NULL, [BIRTHDATE] DATETIME NOT NULL, [PHONE] TEXT(50) NULL, [CELL] TEXT(50) NULL, [EMAIL] TEXT(100) NULL, [PIC] TEXT(255) NULL, [T1] TEXT(255) NULL, [T2] TEXT(255) NULL, [T3] TEXT(255) NULL, [T4] TEXT(255) NULL, [T5] TEXT(255) NULL, [T6] TEXT(255) NULL, [T7] TEXT(255) NULL, [T8] TEXT(255) NULL, [T9] TEXT(255) NULL, [T10] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)
    
    response.write("<h3>create VOLUNTEER table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "VOLUNTEER]([ID] LONG PRIMARY KEY NOT NULL, [FIRSTNAME] TEXT(50) NOT NULL, [LASTNAME] TEXT(50) NOT NULL, [PHONE] TEXT(50) NULL, [CELL] TEXT(50) NULL, [EMAIL] TEXT(100) NULL, [PIC] TEXT(255) NULL, [T1] TEXT(255) NULL, [T2] TEXT(255) NULL, [T3] TEXT(255) NULL, [T4] TEXT(255) NULL, [T5] TEXT(255) NULL, [T6] TEXT(255) NULL, [T7] TEXT(255) NULL, [T8] TEXT(255) NULL, [T9] TEXT(255) NULL, [T10] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create PLAYER_POSITION table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "PLAYER_POSITION]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [POSITION] TEXT(50) NOT NULL, [DESCRIP] MEMO NULL, [SORT] LONG NOT NULL, [TYPE] TEXT(20), [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create PROGRAM table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "PROGRAM]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [PROGRAM] TEXT(50) NOT NULL, [DESCRIP] MEMO NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create ROSTER table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "ROSTER]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [TEAM_ID] LONG NOT NULL, [PLAYER_ID] LONG NOT NULL, [POSITION_ID] LONG NOT NULL, [RANK] LONG NULL, [YEAR] LONG NOT NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create SPONSOR table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "SPONSOR]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [SPONSOR] TEXT(50) NOT NULL, [EMAIL] TEXT(100) NULL, [URL] TEXT(100) NULL, [ADDRESS] TEXT(255) NULL, [PHONE] TEXT(50) NULL, [CELL] TEXT(50) NULL, [FAX] TEXT(50) NULL, [DESCRIP] MEMO NULL, [PIC] TEXT(255) NULL, [T1] TEXT(255) NULL, [T2] TEXT(255) NULL, [T3] TEXT(255) NULL, [T4] TEXT(255) NULL, [T5] TEXT(255) NULL, [T6] TEXT(255) NULL, [T7] TEXT(255) NULL, [T8] TEXT(255) NULL, [T9] TEXT(255) NULL, [T10] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create TEAM table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "TEAM]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [TEAM] TEXT(50) NOT NULL, [DESCRIP] MEMO NULL, [LEAGUE_ID] LONG NULL, [PROGRAM_ID] LONG NOT NULL, [DIVISION_ID] LONG NOT NULL, [SPONSOR_ID] LONG NULL, [COLORS_HOME] TEXT(50) NULL, [COLORS_AWAY] TEXT(50) NULL, [ACTIVE] BYTE NOT NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)

	response.write("<h3>create TEAM_YEARLIES table</h3>")
	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "TEAM_YEARLIES]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [NAME] TEXT(50) NOT NULL, [VALUE] TEXT(255) NULL, [TEAM_ID] LONG NOT NULL, [YEAR] LONG NOT NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
	executeThis(sSql)
    
    reDim arrData(2)
	arrData(0) = STRTABLEPREFIX & "MODS"
	arrData(1) = "[M_CODE],[M_NAME],[M_VALUE]"
    arrData(2) = "'pCount','roster','0'"
	populateB(arrData)

end sub

sub roster_menus()
	mnu_icon = "Themes/<%= strTheme %" & ">/icons/arrow1.gif"
	
	mnu.DelMenuFiles("")
  response.Write("<h3>Forms Menu</h3>")
  sSql = "select APP_ID from PORTAL_APPS where APP_iNAME = '" & app_name & "'"
  set rsT = my_Conn.execute(sSql)
    ap_id = rsT(0)
  set rsT = nothing
  
  'Add roster link to nav menu
  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "APP_ID,INAME,LINK,MNUACCESS,MNUADD,MNUFUNCTION,MNUIMAGE,MNUORDER,MNUTITLE,NAME,ONCLICK,PARENT,PARENTID,TARGET"
  arrData(2) = ap_id & ",'nav_main','roster.asp','','','','',10,'Portal Navbar','Roster','','nav_main',0,'_parent'"
  populateB(arrData)

  'Add roster link to admin managers menu
  sSql = "select ID from menu where Name = 'Managers' and iName = 'b_managers'"
  set rsT = my_Conn.execute(sSql)
  b_managers_pid = rsT(0)
  set rsT = nothing

  redim arrData(2)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,mnuAccess,onClick,mnuImage,mnuTitle,iName,ParentID,app_id,mnuOrder"
  arrData(2) = "'Roster Manager','Managers','admin_roster.asp','_parent','','','','* Managers ADMIN','b_managers'," & b_managers_pid & "," & ap_id & ",15"
  populateB(arrData)
  

end sub


sub roster_skyPage()
	'::::::::::::::::::::: CREATE SPECIAL SKYPAGE :::::::::::::::::::::::::::::
	response.write("<h3>Add Roster to SkyPage Manager</h3>")
	  redim arrData(2)
	  arrData(0) = strTablePrefix & "PAGES"
	  arrData(1) = "P_Name,P_iName,P_TITLE,P_CONTENT,P_ACONTENT,P_LEFTCOL,P_RIGHTCOL,P_MAINTOP,P_MAINBOTTOM,P_APP,P_USE_PG_DISP,P_OTHER_URL,P_CAN_DELETE"
	  arrData(2) = "'Roster Display Page', 'roster','Title is not used','Since this is a special skypage, nothing in this box will ever be visible','','Main Menu:menu_fp','','',''," & app_id & ",0,'roster.asp',0"
	  populateB(arrData)

end sub

sub uninstall_Roster()
	response.Write("<h3>Uninstall Roster</h3>")
	sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = '" & app_name & "'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.EOF then
	apid = rsA("APP_ID")
	end if
	set rsA = nothing
	
	response.write("<h4>Delete Menus</h4>")
	sSql = "delete from menu where APP_ID = " & apid
	executeThis(sSql)
	
	response.write("<h4>Delete SkyPages</h4>")
	sSql = "delete from " & strTablePrefix & "Pages where p_app = " & apid
	executeThis(sSql)
	
	response.write("<h4>Delete From MODS</h4>")
	sSql = "DELETE FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster'"
	executeThis(sSql)
    
    response.write("<h4>Delete From UPLOAD_CONFIG</h4>")
    sSql = "DELETE FROM " & STRTABLEPREFIX & "UPLOAD_CONFIG WHERE [UP_APPID] = " & apid
	
	dropTable(strTablePrefix & "DIVISION")
	dropTable(strTablePrefix & "LEAGUE")
	dropTable(strTablePrefix & "PLAYER")
    dropTable(strTablePrefix & "VOLUNTEER")
	dropTable(strTablePrefix & "PLAYER_POSITION")
	dropTable(strTablePrefix & "PROGRAM")
	dropTable(strTablePrefix & "ROSTER")
	dropTable(strTablePrefix & "SPONSOR")
	dropTable(strTablePrefix & "TEAM")
	dropTable(strTablePrefix & "TEAM_YEARLIES")
	
	response.write("<h4>Delete Roster App</h4>")
	sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_iNAME='" & app_name & "'"
	executeThis(sSql)
	mnu.DelMenuFiles("")

end sub


sub updateRoster(version)
  select case version
    case "1.0"
        'Shouldn't be any upgrades to this, it's first version
        updateVersion version,app_name 'automatically updates our version, cool huh?
    case "1.1"
        'We're changing volunteer and player ID fields to autonumber/identity
        response.write "<h3>Update to v1.1</h3>"
        response.write "<p><b>Backup Player, Volunteer and Roster Information</b></p>"
        
        'Check to see if the bak tables exist
        on error resume next
        err.clear
    	sSQL = "SELECT * INTO [" & STRTABLEPREFIX & "PLAYER_BAK] FROM [" & STRTABLEPREFIX & "PLAYER];"
    	my_Conn.Execute(sSQL)
        
    	sSQL = "SELECT * INTO [" & STRTABLEPREFIX & "VOLUNTEER_BAK] FROM [" & STRTABLEPREFIX & "VOLUNTEER];"
    	my_Conn.Execute(sSQL)
        
        sSQL = "SELECT * INTO [" & STRTABLEPREFIX & "ROSTER_BAK] FROM [" & STRTABLEPREFIX & "ROSTER];"
    	my_Conn.Execute(sSQL)

        err.clear
        on error goto 0
        
        dropTable(STRTABLEPREFIX & "PLAYER")
        dropTable(STRTABLEPREFIX & "VOLUNTEER")
        
    	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "PLAYER]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [FIRSTNAME] TEXT(50) NOT NULL, [LASTNAME] TEXT(50) NOT NULL, [SEX] TEXT(1) NULL, [BIRTHDATE] DATETIME NOT NULL, [PHONE] TEXT(50) NULL, [CELL] TEXT(50) NULL, [EMAIL] TEXT(100) NULL, [PIC] TEXT(255) NULL, [T1] TEXT(255) NULL, [T2] TEXT(255) NULL, [T3] TEXT(255) NULL, [T4] TEXT(255) NULL, [T5] TEXT(255) NULL, [T6] TEXT(255) NULL, [T7] TEXT(255) NULL, [T8] TEXT(255) NULL, [T9] TEXT(255) NULL, [T10] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
    	executeThis(sSql)
        
    	sSQL = "CREATE TABLE [" & STRTABLEPREFIX & "VOLUNTEER]([ID] LONG IDENTITY (1, 1) PRIMARY KEY NOT NULL, [FIRSTNAME] TEXT(50) NOT NULL, [LASTNAME] TEXT(50) NOT NULL, [PHONE] TEXT(50) NULL, [CELL] TEXT(50) NULL, [EMAIL] TEXT(100) NULL, [PIC] TEXT(255) NULL, [T1] TEXT(255) NULL, [T2] TEXT(255) NULL, [T3] TEXT(255) NULL, [T4] TEXT(255) NULL, [T5] TEXT(255) NULL, [T6] TEXT(255) NULL, [T7] TEXT(255) NULL, [T8] TEXT(255) NULL, [T9] TEXT(255) NULL, [T10] TEXT(255) NULL, [AUSER] LONG NOT NULL, [ADATE] DATETIME NOT NULL, [EUSER] LONG NOT NULL, [EDATE] DATETIME NOT NULL);"
    	executeThis(sSql)
        
        response.write "<p><b>Restore Player and Volunteer Information</b></p>"
        restoreCount = 0

        strSql = "SELECT [FIRSTNAME], [LASTNAME], [SEX], [BIRTHDATE], [PHONE], [CELL], [EMAIL], [PIC], [T1], [T2], [T3], [T4], [T5], [T6], [T7], [T8], [T9], [T10], [AUSER], [ADATE], [EUSER], [EDATE] FROM [" & STRTABLEPREFIX & "PLAYER_BAK]"
        set restoreRs = my_conn.execute(strSql)
        
        if restoreRs.EOF or restoreRs.BOF then
            response.write "No player records to restore!<br />"
        else
            while not restoreRs.EOF
                strSql = "INSERT INTO " & STRTABLEPREFIX & "PLAYER ([FIRSTNAME], [LASTNAME], [SEX], [BIRTHDATE], [PHONE], [CELL], [EMAIL], [PIC], [T1], [T2], [T3], [T4], [T5], [T6], [T7], [T8], [T9], [T10], [AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('"
                strSql = strSql & myChkString(restoreRs.Fields("FIRSTNAME"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("LASTNAME"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("SEX"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("BIRTHDATE"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("PHONE"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("CELL"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("EMAIL"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("PIC"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T1"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T2"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T3"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T4"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T5"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T6"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T7"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T8"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T9"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T10"), "sqlstring") & "'"
                strSql = strSql & ", " & myChkString(restoreRs.Fields("AUSER"), "sqlstring") & ""
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("ADATE"), "sqlstring") & "'"
                strSql = strSql & ", " & myChkString(restoreRs.Fields("EUSER"), "sqlstring") & ""
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("EDATE"), "sqlstring") & "');"
                executeThis(strSql)
                
                restoreCount = restoreCount + 1
                restoreRs.movenext
            wend
            
            response.write restoreCount & " player records restored<br />"
        end if
        
        restoreCount = 0

        strSql = "SELECT [FIRSTNAME], [LASTNAME], [PHONE], [CELL], [EMAIL], [PIC], [T1], [T2], [T3], [T4], [T5], [T6], [T7], [T8], [T9], [T10], [AUSER], [ADATE], [EUSER], [EDATE] FROM [" & STRTABLEPREFIX & "VOLUNTEER_BAK]"
        set restoreRs = my_conn.execute(strSql)
        
        if restoreRs.EOF or restoreRs.BOF then
            response.write "No volunteer records to restore!<br />"
        else
            while not restoreRs.EOF
                strSql = "INSERT INTO " & STRTABLEPREFIX & "VOLUNTEER ([FIRSTNAME], [LASTNAME], [PHONE], [CELL], [EMAIL], [PIC], [T1], [T2], [T3], [T4], [T5], [T6], [T7], [T8], [T9], [T10], [AUSER], [ADATE], [EUSER], [EDATE]) VALUES ('"
                strSql = strSql & myChkString(restoreRs.Fields("FIRSTNAME"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("LASTNAME"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("PHONE"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("CELL"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("EMAIL"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("PIC"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T1"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T2"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T3"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T4"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T5"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T6"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T7"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T8"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T9"), "sqlstring") & "'"
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("T10"), "sqlstring") & "'"
                strSql = strSql & ", " & myChkString(restoreRs.Fields("AUSER"), "sqlstring") & ""
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("ADATE"), "sqlstring") & "'"
                strSql = strSql & ", " & myChkString(restoreRs.Fields("EUSER"), "sqlstring") & ""
                strSql = strSql & ", '" & myChkString(restoreRs.Fields("EDATE"), "sqlstring") & "');"
                executeThis(strSql)
                
                restoreCount = restoreCount + 1
                restoreRs.movenext
            wend
            
            response.write restoreCount & " volunteer records restored<br />"
        end if
        
        
        response.write "<p><b>Delete roster data</b></p>"
        
        sSql = "DELETE * FROM [" & STRTABLEPREFIX & "ROSTER]"
        executeThis(sSql)
        
        updateVersion version,app_name 'automatically updates our version, cool huh?
        
    case "1.2"
        'Ability to make volunteer data "private" on a team basis
        response.write "<h3>Update to v1.2</h3>"
        
        strSql = STRTABLEPREFIX & "ROSTER, [PERMS] MEMO NULL"
        alterTable2(strSql)
        
        strSql = "UPDATE " & STRTABLEPREFIX & "ROSTER SET [PERMS] = 0"
        executeThis(strSql)
        
        updateVersion version,app_name 'automatically updates our version, cool huh?

    case else
        response.write "nothing to update"
        
  end select
end sub

sub rosterAddDefaultData()
	response.write("<h3>Adding some default data!</h3>")
	
	redim arrData(4)
	arrData(0) = STRTABLEPREFIX & "LEAGUE"
	arrData(1) = "[LEAGUE],[DESCRIP],[WEBSITE],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'League A','','http://www.leaguea.com'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'League B','','http://www.leagueb.com'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'League C','','http://www.leaguec.com'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)
	
	redim arrData(4)
	arrData(0) = STRTABLEPREFIX & "PROGRAM"
	arrData(1) = "[PROGRAM],[DESCRIP],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'Program A',''," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'Program B',''," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'Program C',''," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)

	redim arrData(4)
	arrData(0) = STRTABLEPREFIX & "DIVISION"
	arrData(1) = "[DIVISION],[DESCRIP],[STARTAGE],[ENDAGE],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'Division A','',15,18," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'Division B','',19,22," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'Division C','',23,26," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)
	
	redim arrData(21)
	arrData(0) = STRTABLEPREFIX & "SPONSOR"
	arrData(1) = "[SPONSOR],[T1],[T2],[T3],[T4],[T5],[T6],[T7],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'Gloria','nunc.ac.mattis@atarcuVestibulum.edu','http://www.google.com','P.O. Box 207, 2006 Sed Rd.','Montgomery','1491425373','6123342516','2039115227'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'Sybil','turpis.non@NullamnislMaecenas.org','http://www.google.com','8414 Aliquam Avenue','Fullerton','6958301827','1219434761','6934384283'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'Lionel','ornare.In@tempus.org','http://www.google.com','522-2408 Taciti Road','Lawton','7442789406','4307150986','4187701153'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(5) = "'Molly','tristique@rutrum.org','http://www.google.com','175 Lorem Rd.','Biddeford','3311561547','9656366520','2583033533'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(6) = "'Linus','montes.nascetur@quamdignissimpharetra.edu','http://www.google.com','Ap #891-5931 Nullam Av.','Murrieta','5118883095','6796021457','1157385071'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(7) = "'Marah','Pellentesque@egestasadui.edu','http://www.google.com','671-841 Orci. Rd.','Tustin','4955278759','9063923749','9781955216'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(8) = "'Candice','dis.parturient@posuereenimnisl.ca','http://www.google.com','Ap #116-7761 Molestie Av.','Mission Viejo','1481627027','8088530602','2792540148'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(9) = "'Piper','faucibus.Morbi@egestasnunc.ca','http://www.google.com','P.O. Box 165, 9962 Nulla Street','Urbana','1876283205','9075477964','5811893803'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(10) = "'Rhona','vitae@diamProin.com','http://www.google.com','Ap #490-5740 Vitae, St.','Monrovia','5219089634','5444372400','6264594149'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(11) = "'Peter','placerat.augue.Sed@etnetuset.com','http://www.google.com','180 Donec Ave','Claremont','5469841242','8782286435','2281017817'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(12) = "'Zelenia','eu.ligula@nuncsed.edu','http://www.google.com','566-140 Erat. Ave','Columbia','3835709061','3424601781','4079076336'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(13) = "'Lane','Proin.sed@nequeNullamut.com','http://www.google.com','632-2595 Dolor. St.','Pocatello','1506523337','4164938293','8538839291'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(14) = "'Yolanda','neque.pellentesque.massa@eu.edu','http://www.google.com','Ap #659-7735 Dictum St.','Wisconsin Rapids','7447729767','3051353263','3339669227'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(15) = "'Colin','turpis@Utsemper.ca','http://www.google.com','Ap #574-3787 Tempor Rd.','Stevens Point','5999431164','6032634906','9226112766'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(16) = "'Lunea','dolor@Sednunc.edu','http://www.google.com','607-7658 Morbi Street','Somerville','6103542501','8389416172','4485577835'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(17) = "'Merrill','nibh.lacinia.orci@adipiscinglobortisrisus.edu','http://www.google.com','2169 Nibh. Road','Moorhead','6535051224','6721298023','7215435807'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(18) = "'Hu','libero.nec@massaIntegervitae.com','http://www.google.com','Ap #663-6880 Et Street','Clairton','4926916254','4688946020','6588920587'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(19) = "'Oleg','Vivamus@vitaerisusDuis.org','http://www.google.com','5869 Integer Rd.','Urbana','3171127158','7076096684','4003873259'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(20) = "'Coby','vitae.purus.gravida@non.com','http://www.google.com','779-5843 Placerat, Ave','Catskill','8379335178','2819552154','9649976342'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(21) = "'Jacob','iaculis.nec@lorem.ca','http://www.google.com','2760 Cum Ave','New Kensington','2172634888','8537413765','2904298991'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)
	
	redim arrData(9)
	arrData(0) = STRTABLEPREFIX & "PLAYER_POSITION"
	arrData(1) = "[POSITION],[DESCRIP],[SORT],[TYPE],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'Manager','',1,'vol'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'Coach','',2,'vol'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'Asst. Coach','',3,'vol'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(5) = "'Trainer','',4,'vol'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(6) = "'Forward','',5,'player'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(7) = "'Left Wing','',6,'player'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(8) = "'Defense','',7,'player'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(9) = "'Goaltender','',8,'player'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)

	redim arrData(81)
	arrData(0) = STRTABLEPREFIX & "PLAYER"
	arrData(1) = "[ID],[FIRSTNAME],[LASTNAME],[BIRTHDATE],[PHONE],[CELL],[PIC],[T1],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "1,'Kitra','Harrington','02/10/1995','9318767562','3078713167','eu','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "2,'Evangeline','Daniels','08/05/1994','1797751683','5692153450','ornare,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "3,'Coby','Mcneil','24/01/1987','5836568647','4136818387','ornare','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(5) = "4,'Flynn','Stephens','13/05/1998','5998214707','3541288602','ipsum.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(6) = "5,'Herman','Benton','13/11/1999','3287215705','6379087342','faucibus','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(7) = "6,'Branden','Faulkner','30/07/1986','3145185313','1357162799','vitae','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(8) = "7,'Rhea','Lambert','23/04/1993','3277880004','5131554689','non','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(9) = "8,'Kylie','Le','10/08/2000','8556530858','6366470984','vestibulum','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(10) = "9,'Wang','Hodge','06/06/1998','1833998743','3964812358','Sed','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(11) = "10,'Linda','Baldwin','08/06/1987','9449943927','3161913200','convallis','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(12) = "11,'Clio','Schneider','07/04/1988','3773633677','6729647864','Morbi','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(13) = "12,'Keely','Fields','28/05/1986','8635552829','2924502174','Phasellus','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(14) = "13,'Stephanie','Tucker','27/11/1996','3532184947','4027856959','Cras','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(15) = "14,'Zenaida','Mejia','27/08/2005','3705183447','6559407800','malesuada.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(16) = "15,'Leila','Conway','13/08/2006','9113219589','9084420076','lacus.','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(17) = "16,'Marvin','Randolph','05/05/2003','2677539440','8617106104','non','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(18) = "17,'Noble','Moses','29/07/1992','5426793960','7246796467','mauris','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(19) = "18,'Rowan','Sosa','28/01/2002','9836402954','5348209923','vulputate','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(20) = "19,'Emmanuel','Clay','24/07/1986','5996775768','3085049574','sollicitudin','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(21) = "20,'Jacob','Gardner','12/07/1988','9775606096','7633301610','eget','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(22) = "21,'Iona','Brock','09/04/1997','8152195062','5322417171','Sed','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(23) = "22,'Dana','Maynard','08/03/1988','7358509314','6341403045','interdum.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(24) = "23,'Moses','Wilder','03/02/1993','1415530276','1361784328','enim','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(25) = "24,'Yoshi','Reyes','05/12/1998','7438298249','7938214351','ligula','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(26) = "25,'Zenaida','Russo','25/12/1988','6261967600','5304553356','Mauris','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(27) = "26,'Scarlett','Meyers','18/01/2000','1374572199','2851461991','Nunc','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(28) = "27,'Olivia','Moody','22/10/1987','9231917054','5117316501','montes,','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(29) = "28,'Sawyer','Meyers','08/03/1988','5888354277','4791845357','elit.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(30) = "29,'Ulla','Mclaughlin','03/06/1998','4979465943','7882169448','ligula','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(31) = "30,'Francis','Carney','26/09/1992','6886781179','1265208987','Suspendisse','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(32) = "31,'Ingrid','Boyer','20/12/1997','4891131710','4985065091','vel,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(33) = "32,'Camden','Knowles','24/01/2002','1923973352','4756476875','metus.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(34) = "33,'Cleo','Stark','19/03/2006','8195210312','5098451837','enim','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(35) = "34,'Kato','Bowman','09/01/1999','6269088898','4194516016','nunc','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(36) = "35,'Desiree','Blackburn','06/10/2005','5589430056','1658547967','lectus','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(37) = "36,'Sean','Workman','02/01/1999','3062692150','2932998953','sed','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(38) = "37,'Price','Mcconnell','27/08/1999','1423512819','9883042136','molestie','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(39) = "38,'Mariam','Bowman','28/01/2001','3232317350','6119939918','eget','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(40) = "39,'Courtney','Kelly','14/04/1991','4516877101','3329680301','lorem,','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(41) = "40,'Jordan','Carter','23/05/2000','9692816278','8666767009','Ut','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(42) = "41,'Rose','Armstrong','19/11/2000','1813864833','1257425392','at','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(43) = "42,'Thomas','Mills','11/11/1994','7967646628','9054897813','in,','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(44) = "43,'Linda','Carter','04/09/1986','9538124820','5842570476','sed','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(45) = "44,'Kelly','Mcknight','08/02/1986','5691252106','9194914599','velit','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(46) = "45,'Magee','Rivers','30/07/1992','1348094553','6898994903','ornare,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(47) = "46,'Judah','Reid','13/06/1993','6463020064','9745208887','Donec','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(48) = "47,'Michelle','Cook','02/04/1989','7751808454','7678642626','odio','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(49) = "48,'Nita','Vaughan','21/12/2000','3561332118','2981445102','et','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(50) = "49,'Rana','Donovan','01/09/1993','4994463720','8531748947','et','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(51) = "50,'Carolyn','Herring','26/06/1989','1279942341','7981601464','quis','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(52) = "51,'Yvonne','Finley','03/09/1986','3492315332','4376627315','est','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(53) = "52,'Marcia','Melton','12/01/2001','3217600850','1816159596','erat.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(54) = "53,'Willow','Spence','30/03/1999','9145900231','9929880861','Donec','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(55) = "54,'Quynn','Hewitt','27/10/1994','4615203228','6197429321','ac,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(56) = "55,'Abraham','Hancock','01/07/2006','1655539900','4217535144','semper','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(57) = "56,'Odette','Sandoval','17/12/1995','3906000065','6295195526','vitae,','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(58) = "57,'Zeus','Romero','30/06/1994','9428396696','3063138757','luctus','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(59) = "58,'Patrick','Saunders','09/02/1987','2054842794','6343930147','Phasellus','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(60) = "59,'Ashely','Dodson','23/09/1989','3605373578','9532326061','rhoncus.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(61) = "60,'Joseph','Gonzalez','30/10/1987','7902414674','2115693556','a','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(62) = "61,'Willa','James','23/09/1996','9553106896','1475158360','ipsum','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(63) = "62,'Aaron','Chang','01/07/1986','8587471337','6582478426','vulputate,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(64) = "63,'Christian','Schmidt','05/01/1992','6885186944','6828265813','Ut','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(65) = "64,'Brenda','Townsend','02/09/1990','9359677356','8744798371','sodales.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(66) = "65,'Paul','Orr','23/03/2003','8808317342','3101466231','auctor','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(67) = "66,'Amela','Mccoy','18/05/1992','4876742532','4741083303','aliquet','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(68) = "67,'Adrian','Zimmerman','08/05/1988','7314310797','4726358623','a,','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(69) = "68,'Charissa','Carey','07/02/2006','7251650481','8238036408','scelerisque','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(70) = "69,'Halla','Vinson','26/08/1997','1628940665','6307798621','tellus.','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(71) = "70,'Kaden','Brady','22/09/1993','7598927099','9839505136','neque','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(72) = "71,'Leroy','Oneal','03/04/1996','7882232992','8955044938','Fusce','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(73) = "72,'Thomas','Hurley','19/05/1988','1208365054','3775677472','velit','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(74) = "73,'Moses','Keith','13/07/1987','7037821852','7883936114','nascetur','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(75) = "74,'Carol','Herring','15/04/2000','3863637331','1242094782','interdum','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(76) = "75,'Maxwell','Montgomery','01/07/2006','8706529806','1798039053','risus.','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(77) = "76,'Dante','Norton','19/05/1997','4094769630','2922738707','mattis.','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(78) = "77,'Bradley','Hurst','11/01/1998','2675201600','9869400693','ultrices,','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(79) = "78,'Marvin','Montgomery','07/08/2000','4649854912','5426026003','scelerisque','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(80) = "79,'Isadora','Bond','18/02/2005','2431006470','5559571923','lectus.','right'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(81) = "80,'Maggie','Flowers','15/08/1989','3764214458','5582061517','Ut','Left'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)
    
	redim arrData(22)
	arrData(0) = STRTABLEPREFIX & "VOLUNTEER"
	arrData(1) = "[ID],[FIRSTNAME],[LASTNAME],[PHONE],[CELL],[PIC],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "81,'Lance','Mcmillan','6793055332','7533090449','sit'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "82,'Kendall','Roy','2656312992','3078400163','faucibus'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "83,'Ella','Jackson','2747395561','2934926123','vestibulum'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(5) = "84,'Jin','Adkins','8133767941','7802363566','mauris.'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(6) = "85,'Thane','Marsh','8197847751','9387238919','dictum'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(7) = "86,'Madonna','Leonard','2211806625','4102863891','Curabitur'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(8) = "87,'Joy','Hendrix','3634299321','5021746157','amet'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(9) = "88,'Gemma','Michael','3378521895','3252573085','vitae'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(10) = "89,'Zena','Saunders','7559086224','1873667892','netus'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(11) = "90,'Amena','Guthrie','5887485140','2498712799','vestibulum'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(12) = "91,'Maggie','Rhodes','5506019698','3478618767','ipsum'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(13) = "92,'Wyoming','Riddle','1961538088','6996594841','Cum'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(14) = "93,'Garrett','Kelley','8924321428','5713801591','ut'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(15) = "94,'Camden','Mcintosh','4571615407','8309828200','Donec'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(16) = "95,'Anne','Reese','4096501969','9205712553','Mauris'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(17) = "96,'Laith','Hayden','6542449958','5097066986','felis'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(18) = "97,'Quinn','Patton','2208651215','4603651161','Etiam'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(19) = "98,'Noel','Hammond','6601277408','3744960320','dui'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(20) = "99,'Margaret','Romero','1288961521','6494902176','amet'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(21) = "100,'Curran','Meyers','1566434463','2594015765','Integer'," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
    arrData(22) = "101,'David','Coates','5555555555','5555555555',''," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
    populateB(arrData)
	
	redim arrData(27)
	arrData(0) = STRTABLEPREFIX & "TEAM"
	arrData(1) = "[TEAM],[LEAGUE_ID],[PROGRAM_ID],[DIVISION_ID],[SPONSOR_ID],[ACTIVE],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "'Team A',1,1,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(3) = "'Team B',1,1,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(4) = "'Team C',1,1,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(5) = "'Team D',1,2,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(6) = "'Team E',1,2,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(7) = "'Team F',1,2,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(8) = "'Team G',1,3,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(9) = "'Team H',1,3,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(10) = "'Team I',1,3,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(11) = "'Team J',2,1,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(12) = "'Team K',2,1,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(13) = "'Team L',2,1,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(14) = "'Team M',2,2,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(15) = "'Team N',2,2,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(16) = "'Team O',2,2,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(17) = "'Team P',2,3,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(18) = "'Team Q',2,3,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(19) = "'Team R',2,3,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(20) = "'Team S',3,1,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(21) = "'Team T',3,1,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(22) = "'Team U',3,1,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(23) = "'Team V',3,2,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(24) = "'Team W',3,2,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(25) = "'Team X',3,2,3," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(26) = "'Team Y',3,3,1," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	arrData(27) = "'Team Z',3,3,2," & randomNum(20) & ",1," & STRUSERMEMBERID & ",'" & NOW() & "'," & STRUSERMEMBERID & ",'" & NOW() & "'"
	populateB(arrData)
	
	reDim arrData(3)
	arrData(0) = STRTABLEPREFIX & "MODS"
	arrData(1) = "[M_CODE],[M_NAME],[M_VALUE]"
	arrData(2) = "'year','roster','2007-2008'"
	arrData(3) = "'year','roster','2006-2007'"
	populateB(arrData)
    
    strSql = "UPDATE " & STRTABLEPREFIX & "MODS SET [M_VALUE] = 101 WHERE [M_NAME] = 'roster' AND [M_CODE] = 'pCount'"
    executeThis(strSql)
	
	strSql = "SELECT [ID] FROM " & STRTABLEPREFIX & "MODS WHERE [M_CODE] = 'year' AND [M_NAME] = 'roster' AND [M_VALUE] = '2007-2008'"
	set rosterRs = my_Conn.execute(strSql)
	year1 = rosterRs.fields("ID")
	set rosterRs = nothing
	
	strSql = "INSERT INTO " & STRTABLEPREFIX & "MODS ([M_CODE],[M_NAME],[M_VALUE]) VALUES ('yearCurrent','roster','" & year1 & "')"
	executeThis(strSql)
	
	strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','http://www.google.com/intl/en_ALL/images/logo.gif',1, " & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
	executeThis(strSql)
	strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','http://l.yimg.com/a/i/ww/beta/y3.gif',1, " & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
	executeThis(strSql)
    strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','http://digg.com/img/feature-meetup.gif',2, " & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
	executeThis(strSql)
	
	reDim arrData(22)
	arrData(0) = STRTABLEPREFIX & "ROSTER"
	arrData(1) = "[TEAM_ID],[PLAYER_ID],[POSITION_ID],[RANK],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]"
	arrData(2) = "1,101,1,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(3) = "1,82,2,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(4) = "1,83,3,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(5) = "1,84,4,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(6) = "1,5,5,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(7) = "1,6,5,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(8) = "1,7,6,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(9) = "1,8,7,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(10) = "1,9,7,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(11) = "1,10,8,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(12) = "1,101,1,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(13) = "1,85,2,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(14) = "1,91,3,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(15) = "1,92,4,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(16) = "1,13,5,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(17) = "1,14,5,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(18) = "1,15,6,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(19) = "1,16,7,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(20) = "1,17,7,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	arrData(21) = "1,18,8,NULL," & year1 + 1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
    arrData(22) = "2,101,1,NULL," & year1 & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "'"
	populateB(arrData)
	

end sub

function myChkString(fString,fField_Type)
    if isBarren(fString) then
        myChkString = fString
    else
        fString = Replace(fString, ";", "&#59;", 1, -1, 1) 
        fString = Replace(fString, "<", "&lt;", 1, -1, 1) 
        fString = Replace(fString, ">", "&gt;", 1, -1, 1) 
        fString = Replace(fString, """", "&quot;", 1, -1, 1) 
        fString = Replace(fString, "'", "&#39;", 1, -1, 1) 
        fString = Replace(fString, "\", "", 1, -1, 1) 
        fString = Replace(fString, "|", "", 1, -1, 1) 
        fString = Replace(fString, "--", "", 1, -1, 1) 
        myChkString = trim(fString)
    end if
end function
%>
