<% 
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2006 Dogg Software All Rights Reserved
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

'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|              Coded by Brandon Williams.             |'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'


dim do_app, app_version, app_id
bUninstall = false

':: leave this as is.
strModTablePrefix = ""
app_version = "1.00"
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
  if incDraftsFp then
    createDrafts()
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the drafts_functions.asp ""include"" file<br>")
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

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::		SUBROUTINES BELOW HERE
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub createDrafts()
  spThemeBlock1_open(intSkin)
  response.Write("<hr><h3>DRAFTS MODULE</h3><br>")
  
  'check if app is existing
  sSql = "SELECT APP_NAME,APP_ID,APP_VERSION FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'drafts'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall then
      uninstall_drafts()
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing

 if not do_app then ':: lets check for upgrade
   select case cur_appVer
     case "1.00"
     'current version
     case "0.95"
     updateVersion app_version,"drafts"
     updateDrafts("1.00")
     case "0.9"
     updateVersion app_version,"drfats"
     updateDrafts("0.95")
     updateDrafts("1.00")
     case "0.8"
	   updateVersion app_version,"drafts" 'automatically updates our version, cool huh?
	   updateDrafts("0.9")
	   updateDrafts("0.95")
       updateDrafts("1.00")
     case "0.5"
	   updateVersion app_version,"drafts" 'automatically updates our version, cool huh?
	   updateDrafts("0.8")
	   updateDrafts("0.9")
	   updateDrafts("0.95")
       updateDrafts("1.00")
   end select
 elseif not bUninstall then
    addApp()
    crMsgTbl()
    b_drafts()

    Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
 end if
 if not bUninstall then
  response.Write("<hr><h3>Drafts Module Installed</h3><br><br>")
 else
  response.Write("<hr><h3>Drafts Module Uninstalled</h3><br><br>")
 end if
  response.Write("<b>Be sure to delete this file (createDrafts.asp) from your server!</b><br><br>")
  response.Write("<a href=""default.asp""><b>Continue</b></a><br><br><br><br>")
  spThemeBlock1_close(intSkin)
end sub
	
sub addApp()
  'create the app
  response.Write("<hr><h4>Update PORTAL_APPS</h4><br>")
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_iNAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_GROUPS_WRITE],[APP_GROUPS_FULL],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_SUBSEC],[APP_CONFIG],[APP_VIEW],[APP_VERSION],[APP_DATE]"
  arrData(2) = "'drafts','drafts',1,0,'2','','',3,3,3,'config_drafts','drafts.asp','" & app_version & "','" & datetostr2(now()) & "'"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'drafts'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub crMsgTbl()
':::::::::::::::::::::::: CREATE DRAFTS TABLE :::::::::::::::::::::::::::::::
  response.Write("<hr><h4>Create DRAFTS table</h4><br>")
sSQL = "CREATE TABLE [" & strTablePrefix & "DRAFTS]([DRAFT_ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [DRAFT_TEXT] MEMO, [DRAFT_ENTRYUSER] LONG NOT NULL, [DRAFT_ENTRYDATE] TEXT(50), [DRAFT_LASTUSER] LONG NOT NULL, [DRAFT_LASTDATE] TEXT(50));"

createTable(checkIt(sSQL))
end sub

sub uninstall_drafts()
  response.Write("<hr><h3>Uninstall App</h3><br>")
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'drafts'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
	apid = rsA("APP_ID")
  end if
  set rsA = nothing
 
  sSql = "delete from menu where APP_ID = " & apid
  executeThis(sSql)
  
  droptable("" & strTablePrefix & "DRAFTS")
	
  sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_iNAME='drafts'"
  executeThis(sSql)
  mnu.DelMenuFiles("")
  response.Write("<h4>Module Uninstall Complete</h4><br><hr><br>")
end sub

':: DRAFTS MENU :::::::::::::::::::::::::::::::::::
sub b_drafts()
	mnu_icon = "Themes/<%= strTheme %" & ">/icons/arrow1.gif"
	
	mnu.DelMenuFiles("")
  response.Write("<hr><h4>Drafts Menu</h4><br>")
  sSql = "select APP_ID from PORTAL_APPS where APP_iNAME = 'drafts'"
  set rsT = my_Conn.execute(sSql)
    ap_id = rsT(0)
  set rsT = nothing

  sSql = "select ID from menu where Name = 'Members' and iName = 'b_members'"
  set rsT = my_Conn.execute(sSql)
  members_pid = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "Menu"
  arrData(1) = "Name,Parent,Link,Target,mnuAccess,onClick,mnuImage,mnuTitle,iName,ParentID,app_id,mnuOrder"
  arrData(2) = "'Drafts', 'Members','drafts.asp','_parent','','','','* Members','b_members'," & members_pid & ","& ap_id &",12"
  arrData(3) = "'Drafts', 'cp_main','drafts.asp','_parent','','','" & mnu_icon & "','Portal Control Panel','cp_main',0,"& ap_id &",6"
  populateB(arrData)

end sub

sub updateDrafts(version)
  select case version
    case "0.8"
      response.write "<h4>Update to v0.8</h4>"
  	case "0.9"
  		response.write "<h4>Update to v0.9</h4>"
  		response.write "<b>Fix App Table</b><br />"
  		
  		strsql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_WRITE = '1,2', APP_VIEW = 'drafts.asp' WHERE APP_ID = " & app_id
  		populateA(strsql)
  		
  		response.write "<hr /><b>Fix Nav Menu</b><br />"
  		
  		strsql = "UPDATE MENU SET MNUACCESS = '', ONCLICK = '' WHERE APP_ID = " & app_id
  		populateA(strsql)
    case "0.95"
      response.write "<h4>Update to v0.95</h4>"
      response.write "<b>Fix App Table</b><br />"
      
   		strsql = "UPDATE " & strTablePrefix & "APPS SET APP_SUBSCRIPTIONS = 3, APP_BOOKMARKS = 3, APP_SUBSEC = 3 WHERE APP_ID = " & app_id
  		populateA(strsql)
        
    case "1.00"
        response.write "<h4>Update to v1.00</h4>"
        response.write "<b>Fix Group Access</b><br />"
        
        strSql = "UPDATE " & strTablePrefix & "APPS SET [APP_GROUPS_USERS] = '2', [APP_GROUPS_WRITE] = '', [APP_GROUPS_FULL] = '' WHERE APP_ID = " & app_id
        populateA(strSql)
        

	end select

end sub
%>
