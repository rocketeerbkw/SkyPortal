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

'/**
' * SkyPortal Forms Module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://www.opensource.org/licenses/mit-license.php MIT License
' */


dim do_app, app_version, app_id
bUninstall = false

':: leave this as is.
strModTablePrefix = ""
app_version = "2.5"
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
  if incFormsFp then
    createForms()
  else
    Response.Write("<p>&nbsp;</p>")
    spThemeBlock1_open(intSkin)
    Response.Write("<p>&nbsp;</p><p>")
    Response.Write("You must add the fp_forms.asp ""include"" file<br>")
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
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" --><%

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::		SUBROUTINES BELOW HERE
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub createForms()
  spThemeBlock1_open(intSkin) %>
  <h1>FORMS MODULE</h1>
	<p>Thank you for choosing the Forms Module.  Credits to David Angell for the <a href="http://www.angells.com/FormCreator" >original version</a>.</p>
	<p>After you install, don't forget to fill in a few of the modules settings (<a href="admin_config_modules.asp" >Admin Options > Managers > Module Manager > Forms</a>) and choose a layout for the forms page (<a href="admin_config_cp.asp" >Admin Options > Managers > Layout Manager > Forms Display Page</a>).</p>
<%
  
  'check if app is existing
  sSql = "SELECT APP_NAME,APP_ID,APP_VERSION FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'forms'"
  set rsA = my_Conn.execute(sSql)
  if not rsA.EOF then
    if bUninstall then
      uninstall_Forms()
	else
      do_app = false
	  app_id = rsA("APP_ID")
	  cur_appVer = rsA("APP_VERSION")
	end if
  end if
  set rsA = nothing

 if not do_app then ':: lets check for upgrade
   select case cur_appVer
	 case "2.5"
	  'current version
     case "2.0"
      updateVersion app_version,"forms"
	  updateForms("2.5")
     case "1.0"
      updateVersion app_version,"forms" 'automatically updates our version, cool huh?
      updateForms("2.0")
     case "0.8"
	   updateVersion app_version,"forms" 'automatically updates our version, cool huh?
	   updateForms("1.0")
	   updateForms("2.0")
   end select
 elseif not bUninstall then
    addApp()
    crTbls()
	skyPage()
    b_Forms()
	'started doing things easy after v2
	'base module created at v2, now we update to current version
	updateForms("2.5")

    Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
 end if
 if not bUninstall then
  response.Write("<h2>Forms Module Installed</h2><br><br>")
 else
  response.Write("<h2>Forms Module Uninstalled</h2><br><br>")
 end if
  response.Write("<p><b>Be sure to delete this file (createForms.asp) from your server!</b></p>")
  response.Write("<p><a href=""default.asp""><b>Continue</b></a></p>")
  spThemeBlock1_close(intSkin)
end sub
	
sub addApp()
  'create the app
  response.Write("<h3>Update PORTAL_APPS</h3>")
  redim arrData(2)
  arrData(0) = "[" & strTablePrefix & "APPS]"
  arrData(1) = "[APP_NAME],[APP_INAME],[APP_ACTIVE],[APP_DEBUG],[APP_GROUPS_USERS],[APP_GROUPS_WRITE],[APP_GROUPS_FULL],[APP_SUBSCRIPTIONS],[APP_BOOKMARKS],[APP_SUBSEC],[APP_CONFIG],[APP_VIEW],[APP_VERSION],[APP_DATE],[APP_TDATA1]"
  arrData(2) = "'forms','forms',1,0,'1,2,3','1,2','1',3,3,0,'config_forms','form.asp','" & app_version & "','" & datetostr2(now()) & "',''"
  populateB(arrData)

  'return app_id
  sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'forms'"
  set rsA = my_Conn.execute(sSql)
    app_id = rsA("APP_ID")
  set rsA = nothing
end sub

sub crTbls()
	':::::::::::::::::::::::: CREATE FORM TABLES :::::::::::::::::::::::::::::::
	response.Write("<h3>Create FORMFIELDS Table</h3>")
	sSQL = "CREATE TABLE [" & strTablePrefix & "FORMFIELDS]("
	sSQL = sSQL & "[ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, "
	sSQL = sSQL & "[FLDLINKFORMID] LONG NOT NULL, "
	sSQL = sSQL & "[FLDCAPTION] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDFIELDTYPE] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDVALIDATION] MEMO NULL, "
	sSQL = sSQL & "[FLDREQUIRED] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDWIDTH] LONG NOT NULL, "
	sSQL = sSQL & "[FLDHEIGHT] LONG NOT NULL, "
	sSQL = sSQL & "[FLDORDER] LONG NOT NULL, "
	sSQL = sSQL & "[FLDDEFAULT] MEMO NULL, "
	sSQL = sSQL & "[FLDOPTIONS] MEMO NULL"
	sSQL = sSQL & ");"
	
	createTable(checkIt(sSQL))
	
	response.write("<h3>Create FORMHEADER Table</h3>")
	sSQL = "CREATE TABLE [" & strTablePrefix & "FORMHEADER]("
	sSQL = sSQL & "[ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, "
	sSQL = sSQL & "[FLDFORMNAME] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDRECIPIENTEMAIL] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDEMAILSUBJECT] MEMO NOT NULL, "
	sSQL = sSQL & "[FLDINTROTEXT] MEMO NULL, "
	sSQL = sSQL & "[FLDTHANKYOU] MEMO NULL, "
	sSQL = sSQL & "[FLDINACTIVETEXT] MEMO NULL, "
	sSQL = sSQL & "[ACTIVE] INT NULL"
	sSQL = sSQL & ");"
	
	createTable(checkIt(sSQL))

end sub

sub skyPage()
	'::::::::::::::::::::: CREATE SPECIAL SKYPAGE :::::::::::::::::::::::::::::
	response.write("<h3>Add Forms to SkyPage Manager</h3>")
	  redim arrData(2)
	  arrData(0) = strTablePrefix & "PAGES"
	  arrData(1) = "P_NAME,P_INAME,P_TITLE,P_CONTENT,P_ACONTENT,P_LEFTCOL,P_RIGHTCOL,P_MAINTOP,P_MAINBOTTOM,P_APP,P_USE_PG_DISP,P_OTHER_URL,P_CAN_DELETE"
	  arrData(2) = "'Forms Display Page', 'form','Title is not used','Since this is a special skypage, nothing in this box will ever be visible','','','','',''," & app_id & ",0,'form.asp',0"
	  populateB(arrData)

end sub

sub uninstall_Forms()
	response.Write("<h3>Uninstall App</h3>")
	sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'forms'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.EOF then
	apid = rsA("APP_ID")
	end if
	set rsA = nothing
	
	sSql = "DELETE FROM MENU WHERE APP_ID = " & apid
	executeThis(sSql)
	
	sSql = "DELETE FROM " & strTablePrefix & "PAGES WHERE P_APP = " & apid
	executeThis(sSql)
	
	droptable("" & strTablePrefix & "FORMFIELDS")
	droptable("" & strTablePrefix & "FORMHEADER")
	
	sSql = "DELETE FROM " & strTablePrefix & "APPS WHERE APP_INAME='forms'"
	executeThis(sSql)
	mnu.DelMenuFiles("")

end sub

':: FORMS MENU :::::::::::::::::::::::::::::::::::
sub b_Forms()
	mnu_icon = "Themes/<%= strTheme %" & ">/icons/arrow1.gif"
	
	mnu.DelMenuFiles("")
  response.Write("<h3>Forms Menu</h3>")
  sSql = "SELECT APP_ID FROM PORTAL_APPS WHERE APP_INAME = 'forms'"
  set rsT = my_Conn.execute(sSql)
    ap_id = rsT(0)
  set rsT = nothing

  sSql = "SELECT ID FROM MENU WHERE NAME = 'Managers' AND INAME = 'b_managers'"
  set rsT = my_Conn.execute(sSql)
  b_managers_pid = rsT(0)
  set rsT = nothing

  redim arrData(3)
  arrData(0) = "MENU"
  arrData(1) = "NAME,PARENT,LINK,TARGET,MNUACCESS,ONCLICK,MNUIMAGE,MNUTITLE,INAME,PARENTID,APP_ID,MNUORDER"
  arrData(2) = "'Forms Manager','Managers','admin_forms.asp','_parent','','','','* Managers ADMIN','b_managers'," & b_managers_pid & "," & ap_id & ",15"
  arrData(3) = "'Forms Manager','b_forms','','','','','','* Forms ADMIN','b_forms',0," & ap_id & ",1"
  populateB(arrData)
  
  sSql = "SELECT ID FROM MENU WHERE NAME = 'Forms Manager' AND INAME = 'b_forms'"
  set rsT = my_Conn.Execute(sSql)
  b_forms_pid = rsT(0)
  set rsT = nothing
  
  redim arrData(6)
  arrData(0) = "MENU"
  arrData(1) = "NAME,PARENT,LINK,TARGET,MNUACCESS,ONCLICK,MNUIMAGE,MNUTITLE,INAME,PARENTID,APP_ID,MNUORDER"
  arrData(2) = "'Home Page','Forms Manager','admin_forms.asp','_parent','','','','* Forms ADMIN','b_forms'," & b_forms_pid & "," & ap_id & ",1"
  arrData(3) = "'Add New Form','Forms Manager','admin_forms.asp?action=NewForm','_parent','','','','* Forms ADMIN','b_forms'," & b_forms_pid & "," & ap_id & ",2"
  arrData(4) = "'Edit Form','Forms Manager','admin_forms.asp?next=EditForm','_parent','','','','* Forms ADMIN','b_forms'," & b_forms_pid & "," & ap_id & ",3"
  arrData(5) = "'Copy Form','Forms Manager','admin_forms.asp?next=CopyForm','_parent','','','','* Forms ADMIN','b_forms'," & b_forms_pid & "," & ap_id & ",4"
  arrData(6) = "'Delete Form','Forms Manager','admin_forms.asp?next=DeleteForm','_parent','','','','* Forms ADMIN','b_forms'," & b_forms_pid & "," & ap_id & ",5"
  populateB(arrData)

end sub

sub updateForms(version)
  select case version
    case "1.0"
      response.write "<h3>Update to v1.0</h3>"
      response.write "<p><b>Fix App Table</b></p>"

      strsql = "UPDATE " & strTablePrefix & "APPS SET APP_GROUPS_WRITE = '1,2', APP_VIEW = 'forms.asp' WHERE APP_ID = " & app_id
      populateA(strsql)

      response.write "<p><b>Fix Nav Menu</b></p>"

      strsql = "UPDATE MENU SET MNUACCESS = '', ONCLICK = '' WHERE APP_ID = " & app_id
      populateA(strsql)

    case "2.0"
      response.write "<h3>Update to v2.0</h3>"
      response.write "<p><b>Update FORMHEADER table<b></p>"
      
      strSql = strTablePrefix & "FORMHEADER"
      strSql = strSql & ",[FLDINACTIVETEXT] MEMO NULL,[ACTIVE] INT NULL"
      alterTable2(checkIt(strSql))
      
      strSql = "UPDATE " & STRTABLEPREFIX & "FORMHEADER SET [FLDINACTIVETEXT] = '<div align=""center"">We&#39;re sorry, this form is no longer available.<br /></div>', [ACTIVE] = 1"
      executethis(checkit(strSql))
      
      strSql = "UPDATE " & STRTABLEPREFIX & "APPS SET [APP_SUBSEC] = 1 WHERE APP_ID = " & app_id
      executeThis(checkit(strSql))
	  
	case "2.5"
		response.write "<h3>Update to v2.5</h3>"
		response.write "<p><b>Update FORMHEADER table</b></p>"
		
		strSql = strTablePrefix & "FORMHEADER"
		strSql = strSql & ",[SENDEMAIL] INT NULL,[SENDPM] INT NULL,[SENDTO] MEMO NULL"
		alterTable2(checkIt(strSql))
		
		strSql = "ALTER TABLE " & strTablePrefix & "FORMHEADER ALTER COLUMN [fldRecipientEmail] MEMO NULL"
		alterTable(strSql)
		
		strSql = "UPDATE " & strTablePrefix & "FORMHEADER SET [SENDEMAIL] = 1, [SENDPM] = 0, [SENDTO] = ''"
		executeThis(checkIt(strSql))
		
	case "3.0"
		response.write "<h3>Update to v3.0</h3>"
		response.write "<p><b>Create FORMRESPONSE table</b></p>"
		
		sSQL = "CREATE TABLE [" & strTablePrefix & "FORMRESPONSE]("
		sSQL = sSQL & "[ID] INT IDENTITY (1, 1) PRIMARY KEY NOT NULL, "
		sSQL = sSQL & "[SESSID] LONG NOT NULL, "
		sSQL = sSQL & "[FORMID] LONG NOT NULL, "
		sSQL = sSQL & "[QUESTIONID] LONG NOT NULL, "
		sSQL = sSQL & "[QUESTIONLANG] MEMO NOT NULL, "
		sSQL = sSQL & "[ANSWERLANG] MEMO NOT NULL"
		sSQL = sSQL & ");"

		createTable(checkIt(sSQL))

      
  end select
end sub
%>
