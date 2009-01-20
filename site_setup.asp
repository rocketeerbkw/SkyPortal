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
dim inFSOenabled, bErr, strVer, longVer, bDBOK, bFPOK
dim fsoMSG, fsoObj, fso, fo, boolPerm, portalUrl, erMsg, tmpMSG

strCharset = "utf-8"
Session.CodePage = "65001"

itName = "sp_IceMan"
itFolder = "sp_IceMan"
itLogo = "Site_Logo.jpg"
itAuthor = "<a href=http://www.frozenwinds.com/>R.Frost</a>"
itDesc = "IceMan Skin Series - IceMan-Pro Sp Version"
itSubSkin = 3
det = ""

'::::::::::::::::::::::::::::::::::::
strVer = "1.0"
longVer = "1.0"
strDebug = false
bDBOK = false
bFPOK = false
'Dim dbHits
dim arrData()
dim indexes()
Dim newTbl, oldTbl, betaTbl, v20Tbl, v21Tbl
Dim vRC1Tbl, vRC2Tbl
'sqlVer = 7
bHasTable = false
sInstallType = ""
bIsUpgrade = false
boolPerm = false
fsoObj = false
fsoMSG = ""
erMsg = ""
tmpMSG = ""
dbHits = 0
newTbl = 0
oldTbl = 0
betaTbl = 0
v20Tbl = 0
v21Tbl = 0
vRC1Tbl = 0
vRC2Tbl = 0
comCode = ""
CustomCode = 0

ErrorCount = Request.QueryString("RC")
comCode = cLng(Request.QueryString("cmd"))
sessCode = session.Contents("setup")
portalUrl = "http://" & request.ServerVariables("SERVER_NAME") & left(Request.ServerVariables("URL"),instrrev(Request.ServerVariables("URL"),"/"))

if ErrorCount = 1 then
  CustomCode = 2
end if

if comCode <> "3" then 'setup
  blnSetup = "Y"
else ' db is created/updated. 
  blnSetup = ""
end if

if comCode = "3" then ' db is created/updated. 
  blnSetup = ""
  resetCoreConfig()
  'Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end if

%><!--#include file="config.asp" -->
<!--#include file="lang/en/core.asp" -->
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="lang/en/core_install_data.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="includes/inc_DBFunctions.asp" -->
<!--#include file="includes/inc_Theme.asp" -->
<%
installTheme = itFolder
if request("debug") = 1 then
  strDebug = true
end if
	
	'check database type
	Select case lcase(strDBType)
		case "access"
			ErrorCount = 0
		case "sqlserver"
			ErrorCount = 0
	'		CustomCode = 1
		case "mysql"
			ErrorCount = 1
	'		CustomCode = 1
		case else
			ErrorCount = 1
			CustomCode = 1
	end Select

if ErrorCount = 0 then
  if comCode = "2" then 
	if strDebug then
	response.Write("<b>comCode = 2, lets get variable info!</b><br /><br />")
	end if
			adminName = request.Form("adminName")
			adminPass = pEncrypt(pEnPrefix & Request.Form("adminPass"))
			siteName = replace(ChkString(request.Form("siteName"),"sqlstring"),"'","''")
			adminEmail = request.Form("adminEmail")
			mailServer = request.Form("mailServer")
			emailAddy = request.Form("emailAddy")
			emailComponent = request.Form("emailComponent")
			instType = ChkString(request.Form("installType"),"sqlstring")
			localhost = ChkString(request.Form("localhost"),"sqlstring")
			if request.Form("timeoffset") <> "" then
			  timeoffset = request.Form("timeoffset")
			else
			  timeoffset = 0
			end if
			
	'response.Write("<b>instType: </b>" & instType & "<br />")
		if instType = "new" or instType = "migrateMwpxNext" then
			dim	buArticle, buDL, buClassified, buForums, buPics, buLinks
			buArticle = cint(request.Form("Articles"))
			buDL = cint(request.Form("Downloads"))
			buClassified = cint(request.Form("Classifieds"))
			buForums = cint(request.Form("Forums"))
			buPics = cint(request.Form("Pictures"))
			buLinks = cint(request.Form("Links"))
		end if
	if strDebug then
	response.Write("Variable listing:<br />")
	response.Write("<b>instType: </b>" & instType & "<br />")
	response.Write("<b>adminName: </b>" & adminName & "<br />")
	response.Write("<b>adminPass: </b>" & adminPass & "<br />")
	response.Write("<b>adminEmail: </b>" & adminEmail & "<br />")
	response.Write("<b>siteName: </b>" & siteName & "<br />")
	response.Write("<b>mailServer: </b>" & mailServer & "<br />")
	response.Write("<b>emailAddy: </b>" & emailAddy & "<br />")
	response.Write("<b>emailComponent: </b>" & emailComponent & "<br />")
	response.Write("<b>localhost: </b>" & localhost & "<br />")
	response.Write("<b>installTheme: </b>" & installTheme & "<br /><br />")
	response.Write("<b>buArticle: </b>" & buArticle & "<br />")
	response.Write("<b>buDL: </b>" & buDL & "<br />")
	response.Write("<b>buClassified: </b>" & buClassified & "<br />")
	response.Write("<b>buForums: </b>" & buForums & "<br />")
	response.Write("<b>buPics: </b>" & buPics & "<br />")
	response.Write("<b>buLinks: </b>" & buLinks & "<br /><br />")
	'response.End()
	end if
  end if
	
	on error resume next
	'check if the connection string will open the database
	'try to open the connection
		set my_Conn = Server.CreateObject("ADODB.Connection")
		my_Conn.Open strConnString	
	
		'if there is an error,  show error box
		for counter = 0 to my_conn.Errors.Count -1
			ConnErrorNumber = my_conn.Errors(counter).Number
			ConnErrorDesc = my_conn.Errors(counter).Description
			if ConnErrorNumber <> 0 and ConnErrorNumber <> -2147217887 then 
			    writeToLog "Database","",ConnErrorNumber & " : " & ConnErrorDesc
				ErrorCount = 1
				CustomCode = 2
				ErrorCode = ConnErrorNumber & "<br />" & ConnErrorDesc
				my_conn.Errors.Clear 
			end if
		next
		my_Conn.Errors.Clear
		Err.Clear
	
	' debugging
	if strDebug then
	response.Write("Start FSO check<br />")
	end if
	'check for FileSystemObject
	fsoObj = fsoCheck()
	' debugging
	if strDebug then
	response.Write("End FSO check: " & fsoObj & "<br /><br />")
	end if
	
	'response.Write("<br />bHasTable0:" & bHasTable)
	'test for v1.3x member table
	' if this table is missing, it is a new install
	strSql = "SELECT MEMBER_ID FROM PORTAL_MEMBERS"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	my_Conn.Errors.Clear
	Err.Clear
	
  if bHasTable then
  ':: test for MWPX database
	'strSql = "SELECT COL_ID FROM PORTAL_LAYOUT"
	'my_Conn.Execute strSql
	'Call CheckSqlError("mwpx_next")
	'my_Conn.Errors.Clear
	'Err.Clear
	'response.Write("<br />bHasTable1:" & bHasTable)
    'if bHasTable then
	  'bHasTable = false
    'end if
  end if
	
	'response.Write("<br />bHasTable1:" & bHasTable)
  if bHasTable then
	'Lets test for v1.5 new table field
	'this field is added in v1.5 - not in v1.3x
	strSql = "SELECT B_NAME FROM PORTAL_BANNERS"
	my_Conn.Execute strSql
	Call CheckSqlError("v13x")
	my_Conn.Errors.Clear
	Err.Clear
	
	'response.Write("<br />bHasTable2:" & bHasTable)
	if bHasTable and sInstallType = "v13x" then
	'Lets test for v1.5 default URL in case the folder was renamed
	strSql = "SELECT C_STRHOMEURL FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	'Insert new values
	strSql = "UPDATE PORTAL_CONFIG SET C_STRHOMEURL='" & portalUrl & "' WHERE CONFIG_ID = 1"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	my_Conn.Errors.Clear
	Err.Clear
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
	end if
  end if
	
	'response.Write("<br />bHasTable3:" & bHasTable)
  if bHasTable then
	'Lets test for v1.5b3 new table field
	'this field is added in v1.5b3 - not in v1.5
	' upgrades to v2.0 either started with v1.3x OR v1.5b3
	strSql = "SELECT THEME_ID FROM PORTAL_MEMBERS"
	my_Conn.Execute strSql
	Call CheckSqlError("v15b3")
	my_Conn.Errors.Clear
	Err.Clear
  end if
	'response.Write("<br />bHasTable4:" & bHasTable)
  if bHasTable then
	'Lets test for v2.1 new table field
	'this field is added in v2.1 - is not in v2.0
	strSql = "SELECT C_SECIMAGE FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("v20")
	my_Conn.Errors.Clear
	Err.Clear
  end if
	
	'response.Write("<br />bHasTable5:" & bHasTable)
  if bHasTable then
	'Lets test for SP RC1 new table field
	'this field is added in SP RC1 - is not in v2.1x
	strSql = "SELECT C_INTSUBSKIN FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("v21")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  if bHasTable then
	'Lets test for SP RC2 new table field
	'this field is added in SP RC2 - is not in RC1
	strSql = "SELECT APP_GROUPS_FULL FROM PORTAL_APPS"
	my_Conn.Execute strSql
	Call CheckSqlError("vRC1")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  if bHasTable then
	'Lets test for SP RC3 new table field
	'this field is added in SP RC3 - is not in RC2
	strSql = "SELECT id FROM Menu"
	my_Conn.Execute strSql
	Call CheckSqlError("vRC2")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  if bHasTable then
	'We now start to track SP by version # from the db.
	strSql = "SELECT C_PORTAL_VERSION FROM PORTAL_CONFIG WHERE CONFIG_ID = 1"
	set rsSp = my_Conn.Execute(strSql)
	if rsSp.eof then
	else
	  sInstallType = rsSp("C_PORTAL_VERSION")
	  select case sInstallType
	    case "RC7"
		  sInstallType = "RC7"
		  bHasTable = false
		  session.Contents("setup") = "sp_070407sd"
	    case "RC6"
		  sInstallType = "RC6"
		  bHasTable = false
		  session.Contents("setup") = "sp_070407sd"
	    case "RC5"
		  sInstallType = "RC5"
		  bHasTable = false
		  session.Contents("setup") = "sp_070407sd"
	    case "RC4"
		  sInstallType = "RC4"
		  bHasTable = false
		  session.Contents("setup") = "sp_070407sd"
	    case "RC3"
		  vRC3Tbl = 1
		  bHasTable = false
		  session.Contents("setup") = "sp_070407sd"
		case else
	  end select
	end if
	set rsSp = nothing
	'Call CheckSqlError("vRC2")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  		
	'response.Write("<br />bHasTable6:" & bHasTable)			
	
	on error goto 0
	
end if 'responsecode = 0
Response.Buffer = True
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html> 
<head> 
<!-- This page is generated by SkyPortal / SkyPortal.net <%= date() %> -->
<title>SkyPortal v<%= strVer %> | Site Setup</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=<%= strCharset %>">
<META HTTP-EQUIV="imagetoolbar" CONTENT="no">
<META NAME="AUTHOR" CONTENT="SkyPortal www.SkyPortal.net">
<META NAME="GENERATOR" CONTENT="SkyPortal - http://www.SkyPortal.net">
<meta name="COPYRIGHT" content="Portal code is Copyright (C)2005 - 2008 Tom Nance All Rights Reserved">
<meta http-equiv="Content-Style-Type" content="text/css">
<link rel="stylesheet" href="Themes/<%= installTheme %>/style_core.css" type="text/css">
<script type="text/javascript" src="includes/scripts/prototype.js"></script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<script type="text/JavaScript">
function showMe(id) {
	var o1=document.getElementById(id);
		o1.style.display="block";
}
function hideMe(id) {
	var o1=document.getElementById(id);
		o1.style.display="none";
}
function serverOK(){
  window.location="site_setup.asp"
}

function popHelp(hlp){
	switch (hlp) {
		case "timeoffset":
			alert('<%= jshTimeOffset %>');
			break;
		case "adminEmail":
			alert('<%= jshAdminEmail %>');
			break;
		case "adminpass":
			alert('<%= jshAdminPass %>');
			break;
		case "siteEmail":
			alert('<%= jshSiteEmail %>');
			break;
		default:
			
	}
}

function chkInst(){
  var e = 0;
  siteName = $F('siteName');
  adminEmail = $F('adminEmail');
  mailServer = $F('mailServer');
  emailAddy = $F('emailAddy');
  if (siteName == ''){
    alert("siteName");
	e++;
  }
  if (adminEmail == ''){
    alert("adminEmail");
	e++;
  }
  if (mailServer == ''){
    alert("mailServer");
	e++;
  }
  if (emailAddy == ''){
    alert("emailAddy");
	e++;
  }
  alert(e);
  return false;
}
</script>
<form action="site_setup.asp?cmd=2" method="post" id="form2" name="form2">
<table class="spThemePage" width="100%" align="center" border="0" cellpadding="0" cellspacing="0"><tr><td>

<a name="top"></a>
<% spThemeHeader_open() %>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="300" align="left" valign="middle"><img alt="SkyPortal v<%= strVer %>" border="0" src="Themes/<%= installTheme %>/site_Logo.jpg"></td>
    <td align="right"><img src="files/banners/webdogg.jpg"></td>
  </tr>
</table>
<% spThemeHeader_close() %>

<table class="spThemePage" border="0" width="100%" align="center">
<tr>
<td class="leftPgCol"></td>
<td valign="top" align="center" class="mainPgCol">
  <% 
  
  spThemeTitle = "SkyPortal v" & strVer & "&nbsp;" & txtSUSiteSetup & ""
  spThemeBlock1_open(1)
  %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0">
<tr><td bgcolor="#E7E7EA" align="center" valign="top">
<% 
  	
	if strDebug and comCode = "2" then
	  response.Write("check if sessCode matches..<br />")
	  response.Write("sessCode: " & sessCode & "<br />")
	end if
if comCode = "2" and sessCode = "sp_070407sd" and session.Contents("chkServer") = "OK" then
	if strDebug then
	 response.Write("comCode = 2, sessCode matches..<br /><br />")
	end if
	'instType = chkString(request.Form("installType"),"sqlstring")
	  if strDebug then
	  response.Write("Select install type...<br />")
	  end if
	select case instType
		case "new"
	  		if strDebug then
	  		response.Write("Install type: NEW...<br />")
	  		end if
			if request.Form("adminPass") <> request.Form("adminPass2") then
				session.Contents("erMsg") = "<font color=""#FF0000""><b>Your passwords didn't match</b></font>"
				showInstall()
			else
				createDB()
			end if
		case "upgrade13"
	  		if strDebug then
	  		response.Write("Install type: upgrade13...<br />")
	  		end if
			update13x()
			update20x21()
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgrade15b3"
	  		if strDebug then
	  		response.Write("Install type: upgrade15b3...<br />")
	  		end if
			update15b3()
			update20x21()
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgrade20"
	  		if strDebug then
	  		response.Write("Install type: upgrade20...<br />")
	  		end if
			update20x21() 
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "migrateMwpxNext"
	  		if strDebug then
	  		response.Write("Install type: Migrate from MWPX Next...<br />")
	  		end if
			migrate_mwpx()
			update_211xRC1()
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgrade21"
	  		if strDebug then
	  		response.Write("Install type: upgrade 21...<br />")
	  		end if
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgradeRC1"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC1...<br />")
	  		end if
			update_rc1_rc2()
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgradeRC2"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC2...<br />")
	  		end if
			update_rc2_rc3()
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgradeRC3"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC3...<br />")
	  		end if
			update_rc3_rc4()
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
			update_rc7_v1()
		case "upgradeRC4"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC4...<br />")
	  		end if
			update_rc4_rc5()
			update_rc5_rc6()
			update_rc6_rc7()
  			rc7_clean_members()
			update_rc7_v1()
		case "upgradeRC5"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC5...<br />")
	  		end if
			update_rc5_rc6()
			update_rc6_rc7()
  			rc7_clean_members()
			update_rc7_v1()
		case "upgradeRC6"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC6...<br />")
	  		end if
			update_rc6_rc7()
  			rc7_clean_members()
			update_rc7_v1()
		case "upgradeRC7"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC7...<br />")
	  		end if
			update_rc7_v1()
		case else
			Response.Write("Install Type undefined")
			session.Contents("setup") = ""
			session.Contents("chkServer") = ""
			response.End()
	end select
	 	if ErrorCount = 0 then
		  mnu.DelMenuFiles("")
	  	  response.Write("<b>" & txtSUInstComp & "</b><br />")
		  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
		else
	  	  response.Write("<font color=""#FF0000""><b>" & txtSUInstCompErr & "</b></font><br /><br />")
	  	  response.Write("<br /><b>ErrorCount: " & ErrorCount & "</b><br /><br />")
	  	end if
		session.Contents("setup") = ""
	if ErrorCount = 0 then
	 if not strDebug then
	  response.write("<br /><br /><a href=""site_setup.asp?cmd=3""><h4>" & txtSUContSetup & "</h4></a>")
	  response.Write "<br /><br />" & dbHits & " " & txtSUDBHits & "<br /><br />"
	 else
	  response.Redirect("site_setup.asp?cmd=3")
	 end if
	else
	  'response.write("<br /><br /><a href=""site_setup.asp?cmd=3""><h4>" & txtSUContSetup & "</h4></a>")
	end if
else
		'if sInstallType <> "" then
		'response.Write("sInstallType: " & sInstallType)
		'response.Write("<br />v21Tbl: " & v21Tbl)
		'response.Write("<br />v20Tbl: " & v20Tbl)
		'response.Write("<br />oldTbl: " & oldTbl)
		'response.Write("<br />newTbl: " & newTbl)
		'response.Write("<br />betaTbl: " & betaTbl)
		'response.Write("<br />sInstallType: " & sInstallType)
		'response.Write("<br />comCode: " & comCode)
		'response.Write("<br />bHasTable: " & bHasTable)
	if ErrorCount = 0 then
		if bHasTable then
		  'has all tables
			Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
			showInstalled()
			session.Contents("setup") = ""
		else
		  if session.Contents("chkServer") <> "OK" then
		    if bFso then
			  'checkServer()
			  session.Contents("chkServer") = "OK"
			  response.Redirect("site_setup.asp")
			else
			  session.Contents("chkServer") = "OK"
			  response.Redirect("site_setup.asp")
			end if
		  else
			showInstall()
		  end if
		end if 
	else
		errDisplay()
	end if
end if
%>
</td></tr></table>
<% spThemeBlock1_close(intSkin) %>
</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" align="center" width="100%">
<tr>
<td align="left" class="sp_FootLeft"></td>
<td align="left" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1"><a href="privacy.asp">Privacy</a></font></td>
<td align="right" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1">© 2005-2008 SkyPortal.net&nbsp;<%= txtSUAllRtsReserved %>.</font></td>
<td align="right" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1"><a href="http://www.SkyPortal.net" title="Powered By: SkyPortal.net Version <%= strVer %>" target="_blank">SkyPortal.net</a></font></td>
<td width="20" class="sp_FootTile" nowrap></td>
<td class="sp_FootRight"></td></tr>
</table>
</td>
</tr>
</table></form>
</td></tr></table>
</body>
</html>
<%
sub CheckSqlError(typ)
		'response.Write("<br />CheckSqlError: " & typ)
		'response.Write("<br />Errors.Count: " & my_Conn.Errors.Count)
  dim ChkConnErrorNumber
  if my_Conn.Errors.Count <> 0 or Err.number > 0 then  
	  for counter = 0 to my_Conn.Errors.Count '-1
		ChkConnErrorNumber = my_Conn.Errors(counter).Number
			my_Conn.Errors.Clear 
			Err.Clear
			session.Contents("setup") = "sp_070407sd"
		  	bHasTable = false
			select case typ
			  case "new"
				newTbl = 1
				sInstallType = typ
			  case "v13x"
				oldTbl = 1
				sInstallType = typ
			  case "v15b3"
				betaTbl = 1
				sInstallType = typ
			  case "v20"
				v20Tbl = 1
				sInstallType = typ
			  case "v21"
				v21Tbl = 1
				sInstallType = typ
			  case "vRC1"
				vRC1Tbl = 1
				sInstallType = typ
			  case "vRC2"
				vRC2Tbl = 1
				sInstallType = typ
			  case "mwpx_next"
			    mwpx_next = 1
			    sInstallType = typ
				bHasTable = true
			end select
		next
		  'bHasTable = false
		else
		  select case typ
			case "mwpx_next"
			  if comCode = "3" then
		        bHasTable = true
			  else
		        bHasTable = false
			    session.Contents("setup") = "sp_070407sd"
			  end if
			  mwpx_next = 1
			  sInstallType = typ
			case else
		      bHasTable = true
	  	  end select
		end if
end sub

Sub showInstalled()
    session.Contents("chkServer") = ""
	 Err.Clear
 %>
<!--meta http-equiv="Refresh" content="3; URL=default.asp"--><br /><br /><br />
<table border="1" bgColor="#EAFFFF" cellspacing="0" cellpadding="5" width="80%" height="50%" align="center" bordercolor="#FFFFFF">
	<tr>
		<td align="center">
		<p>
		<font face="Verdana, Arial, Helvetica" size="3">
		<b><%= txtSUCongrats %></b><br /><br />
		<%= txtSUSetupComp %><br /><br /></font>
	 <% 'checkServer() %><br />
		<font face="Verdana, Arial, Helvetica" size="2">
		<b><%= txtSUChkGenSet %></b></font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="Verdana, Arial, Helvetica" size="2">
		<a href="default.asp" target="_top"><%= txtSUContinue %>&nbsp;>>></a>
		</font></td>
	</tr>
</table><br /><br />
<% end sub

sub shoModules() %>
  <p><b><%= txtSUUpgMods %></b><br />
  <%= txtSUUpdMods2 %></p>
  <fieldset style="width:200px;padding:5px;"><legend><%= txtSUMods %></legend>
  <table border="0" cellPadding="2" cellSpacing="0">
    <tr> 
      <td width="20%" align="right" vAlign="top"> 
        <input type="checkbox" name="Articles" value="1" checked>&nbsp;
      </td>
      <td width="80%"><%= txtArticles %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Classifieds" value="1" checked>&nbsp;
      </td>
      <td><%= txtClassifieds %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Downloads" value="1" checked>&nbsp;
      </td>
      <td><%= txtDownloads %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Forums" value="1" checked>&nbsp;
      </td>
      <td><%= txtForums %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Links" value="1" checked>&nbsp;
      </td>
      <td><%= txtLinks %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Pictures" value="1" checked>&nbsp;
      </td>
      <td><%= txtPics %></td>
    </tr>
  </table></fieldset>
<%
end sub

Sub showInstall()
  'session.Contents("chkServer") = "" %>
  <table width="100%" border="1" cellPadding="8" cellSpacing="0" bordercolor="#FFFFFF">
	<tbody>
    <tr align="center"> 
      <td vAlign="top">
	    <% If session.Contents("erMsg") <> "" Then %>
		<br /><%= session.Contents("erMsg") %><br /><br />
		<% End If %>
		<% If sInstallType = "new" Then %>
		<% Else
		     If sInstallType = "mwpx_next" Then %>
			  <h4><%= replace(txtMigrFrom,"[%app%]","MWPX Next") %>?</h4>
			  <% 
			 else %>
			  <h4><%= txtSUUpgradeTo %>&nbsp; v<%= strVer %>?</h4>
			  <% 
			 end if
			  select case sInstallType
				'case "new"
				case "v13x"%>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;MWP v1.3x") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgrade13"><%
				case "v15b3" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;MWP.info v1.5 beta3") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgrade15b3"><%
				case "v20" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;MWP.info v2.0") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgrade20"><%
				case "v21" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;MWP.info v2.1") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgrade21">
				  <%
				case "vRC1" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal vRC1") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC1">
				  <%
				case "vRC2" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal vRC2") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC2">
				  <%
				case "RC3" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal RC3") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC3">
				  <%
				case "RC4" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal RC4") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC4">
				  <%
				case "RC5" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal RC5") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC5">
				  <%
				case "RC6" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal RC6") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC6">
				  <%
				case "RC7" %>
                  <font color="#0000FF"><b>
				  <%= replace(txtUpgrFrom,"[%ver%]","&nbsp;SkyPortal RC7") %>
				  </b></font>
                  <input type="hidden" name="installType" value="upgradeRC7">
				  <%
				case "mwpx_next" %>
                  <font color="#0000FF"><b>
				  <% 'replace(txtMigrFrom,"[%app%]","MWPX Next") %><br />
	You are about to migrate your current MWPX database to the SkyPortal database structure.<br />
	During the migration, there will be changes made to the structure of your MWPX database in<br />
	order to make it compatable with the SkyPortal software. Please make sure that you have a<br />
	current backup of this database that you are about to migrate before proceeding..
	</b></font>
				  <% 
				  If strDBType = "access" then
				    'Response.Write("<br /><br />" & txtDb2Migrate)
				  
				  end if %>
                  <input type="hidden" name="installType" value="migrateMwpxNext">
				  <% 
				case else
				  response.Write("<br />install type: <strong>" & sInstallType & "</strong></br>")
			  end select
			  %>
    <table width="100%" border="0" cellPadding="8" cellSpacing="0">
      <tr align="center"> 
      	<td height="15"></td>
	  </tr>
	</table>
		<% End If %>
	  </td></tr>
	</tbody>
  </table>
<% If sInstallType <> "new" Then
     If sInstallType="v13x" or sInstallType="v15b3" or sInstallType="v20" or sInstallType="v21" or sInstallType="mwpx_next" Then
		shoModules()
	 end if %>
	<p><%= txtSUClkBtn %></p>
	<table width="100%" border="1" cellPadding="8" cellSpacing="0" bordercolor="#FFFFFF">
      <tr> 
        <td colspan="2" align="center"> 
		  <input type="hidden" name="localhost" value="0">
          <input class="button" type="submit" name="Submit" value="<%= txtSUInstSkyPortal %>&nbsp;v<%= strVer %>">
        </td>
	  </tr>
	</table>
<% Else
	arMailServer = split(request.ServerVariables("SERVER_NAME"),".")
	sMailServer = "mail"
	sPortalEmail = "mail@"
	if ubound(arMailServer) < 2 then
	  sMailServer = sMailServer & "." & request.ServerVariables("SERVER_NAME")
	  sPortalEmail = sPortalEmail & request.ServerVariables("SERVER_NAME")
	else
	  for a = 1 to ubound(arMailServer)
	    sMailServer = sMailServer & "." & arMailServer(a)
	    sPortalEmail = sPortalEmail & arMailServer(a) & "."
	  next
	  sPortalEmail = left(sPortalEmail,len(sPortalEmail)-1)
	end if 
	%>
	<div id="newInstall" style="display:block;">
    <table width="100%" border="1" cellPadding=8 cellSpacing=0 bordercolor="#FFFFFF" bgcolor="#F1F1F4">
      <tr> 
      	<td colspan="2" align="center">
			  <span class="fSubTitle"><%= txtSUGetStart %></span><br /><br />
			  <strong><%= txtSUFillOutFrm %></strong>
        	  <!-- <font color="#0000FF"><b><%= txtSUNewInst %></b></font><br /><br /> -->
              <input type="hidden" name="installType" value="new">
			  <hr />
		</td></tr>
      <tr>
        <td align="right">Portal URL: </td>
		<td><strong><%= request.ServerVariables("SERVER_NAME") & Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/")) %></strong><!-- <input class="textbox" name="siteURL" type="text" value=""> --></td>
	  </tr>
      <tr>
        <td align="right"><%= txtSUSiteName %>: </td>
		<td><input class="textbox" id="siteName" name="siteName" type="text" value="<%= txtSUMySite %>"></td>
	  </tr>
      <tr> 
        <td colspan="2" align="center">
		<%= txtSUSameNameAs %><br /><br />
		<%= txtSUSAName %>: <b><%= left(strWebMaster, instr(strWebMaster,",")-1) %></b>
        <input class="textbox" name="adminName" type="hidden" id="adminName" value="<%= left(strWebMaster, instr(strWebMaster,",")-1) %>">
        </td>
      </tr>
      <tr> 
        <td align="right"><%= txtSUDefPass %></td>
		<td><input class="textbox" name="adminPass" type="password" id="adminPass" value="<%= txtSUPassAdmin %>">&nbsp;<img src="themes/<%= itFolder %>/icons/help.gif" onclick="popHelp('adminpass')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;">
        </td>
      </tr>
      <tr> 
        <td align="right"><%= txtSUEnterPassAgin %></td>
		<td><input class="textbox" name="adminPass2" type="password" id="adminPass2" value="<%= txtSUPassAdmin %>"></td>
      </tr>
      <tr>
        <td align="right"><%= txtSUEnterAdmEml %></td>
		<td><input class="textbox" name="adminEmail" type="text" id="adminEmail" value="<%= sPortalEmail %>">&nbsp;<img src="themes/<%= itFolder %>/icons/help.gif" onclick="popHelp('adminEmail')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
      </tr>
      <tr> 
        <td align="right"><%= txtSUSrvrTime %></td>
		<td><%= now() %></td>
      </tr>
      <tr> 
        <td align="right"><%= txtSUTimeOffset %></td>
		<td>
		<select name="timeoffset" id="timeoffset">
		<% 
		for xx = -12 to 12
		  if xx = 0 then
		    response.Write("<option value=""" & xx & """ selected=""selected"">" & xx & "</option>" & vbcrlf)
		  else
		    response.Write("<option value=""" & xx & """>" & xx & "</option>" & vbcrlf)
		  end if		
		next %>
		</select>&nbsp;<img src="themes/<%= itFolder %>/icons/help.gif" onclick="popHelp('timeoffset')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
      </tr>
      <tr>
        <td colspan="2" align="center"><hr /><%= txtSUDetEmlComp %></td>
      </tr>
      <tr>
        <td align="right"><%= txtSUSelEmlComp %>: </td>
		<td><% getEmailComponents() %></td>
      </tr>
      <tr>
	    <td align="right"><%= txtSUEmlServer %>: </td>
		<td><input class="textbox" name="mailServer" type="text" value="<%= sMailServer %>"></td>
      </tr>
      <tr> 
        <td align="right"><%= txtSUSiteEmlAdd %>: </td>
		<td><input class="textbox" name="emailAddy" type="text" value="<%= sPortalEmail %>">&nbsp;<img src="themes/<%= itFolder %>/icons/help.gif" onclick="popHelp('siteEmail')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
   	  </tr>
      <tr>
        <td colspan="2" align="center"><hr /></td>
      </tr>
      <tr> 
        <td colspan="2" align="center"> 
		  <input type="hidden" name="localhost" value="0">
          <!-- <input class="button" type="button" name="Submit" onClick="chkInst();" value="<%= txtSUInstSkyPortal %>&nbsp;v<%= strVer %>"> -->
          <input class="button" type="submit" name="Submit" value="<%= txtSUInstSkyPortal %>&nbsp;v<%= strVer %>">
        </td>
	  </tr>
	</table></div>
<% End If  ':: end sInstallType
end sub

sub errDisplay()
%>
<table border="1" cellspacing="0" cellpadding="5" width="95%" align="center" bordercolor="#FFFFFF">
	<tr>
		<td bgColor=pink align="center">
		<font face="Verdana, Arial, Helvetica" size="4"><%= txtSUThereIsError %></font>
		<p>
		<font face="Verdana, Arial, Helvetica" size="2">
<%
Select case CustomCode
	 case 1 %>
		<%= txtSuCC1 %><br />	<br />
<% case 2 %>
		<%= txtSuCC2 %><br /><br />
<% case 3 %>
		<%= txtSuCC3 %><br /><br />
<% case 4 %>
		<%= txtSuCC4 %><br /><br />
<% case else %>
		<%= txtSuCC5 %><br />
		<br />
<%
end select

		if ErrorCode <> "" then 
			Response.Write("</p><p>" & txtSUErrCode & " :  " & ErrorCode & " ")
			Response.Write("</p><p>" & strDBPath & "</p>")
		end if
%>
		</font></p></td>
	</tr>
	<tr>
		<td align="center"><font face="Verdana, Arial, Helvetica" size="2"><a href="site_setup.asp" target="_top"><%= txtSUClikToRetry %></a></font></td>
	</tr>
</table>
<%
End sub

function fsoCheck()
	 dim tFso
	 tFso = false
     on error resume next
     err.clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   tFso = true
	 end if
	 set fso = nothing
     on error goto 0
	 fsoCheck = tFso
end function

function getEmailComponents()
Dim arrComponent(10)
Dim arrValue(10)
Dim arrName(10)

' components
arrComponent(0) = "CDO.Message"
arrComponent(1) = "CDONTS.NewMail"
arrComponent(2) = "SMTPsvg.Mailer"
arrComponent(3) = "Persits.MailSender"
arrComponent(4) = "SMTPsvg.Mailer"
arrComponent(5) = "CDONTS.NewMail"
arrComponent(6) = "dkQmail.Qmail"
arrComponent(7) = "Geocel.Mailer"
arrComponent(8) = "iismail.iismail.1"
arrComponent(9) = "Jmail.smtpmail"
arrComponent(10) = "SoftArtisans.SMTPMail"

' component values
arrValue(0) = "cdosys"
arrValue(1) = "cdonts"
arrValue(2) = "aspmail"
arrValue(3) = "aspemail"
arrValue(4) = "aspqmail"
arrValue(5) = "chilicdonts"
arrValue(6) = "dkqmail"
arrValue(7) = "geocel"
arrValue(8) = "iismail"
arrValue(9) = "jmail"
arrValue(10) = "smtp"

' component names
arrName(0) = "CDOSYS (IIS 5/5.1/6)"
arrName(1) = "CDONTS (IIS 3/4/5)"
arrName(2) = "ASPMail"		'yes
arrName(3) = "ASPEMail"	'yes
arrName(4) = "ASPQMail"	'yes			'
arrName(5) = "Chili!Mail (Chili!Soft ASP)"	'
arrName(6) = "dkQMail"						'
arrName(7) = "GeoCel"						'
arrName(8) = "IISMail"					'
arrName(9) = "JMail"						'
arrName(10) = "SA-Smtp Mail"

'Dim i
'for i=0 to UBound(arrComponent)
'	if isInstalled(arrComponent(i)) then
'	end if
'next

Response.Write("<select name=""emailComponent"">") & vbcrlf
'Response.Write("<ul>") & vbcrlf
'Response.Write("<option value=""none"" selected="selected"></option>") & vbcrlf
Dim i
for i=0 to UBound(arrComponent)
	if isInstalled(arrComponent(i)) then
	  'Response.Write("<li>"  & arrName(i) &"</li>") & vbcrlf
	  Response.Write("<option value=""" & arrValue(i) & """>" & arrName(i) &"</option>") & vbcrlf
	end if
next
'Response.Write("</ul>") & vbcrlf
Response.Write("</select>") & vbcrlf
end function				'

Function isInstalled(obj)
	on error resume next
	installed = False
	Err = 0
	Dim chkObj
	Set chkObj = Server.CreateObject(obj)
	If 0 = Err Then installed = True
	Set chkObj = Nothing
	isInstalled = installed
	Err = 0
	on error goto 0
End Function

 %>
<!--#include file="install/createUpgrade.asp" -->
<!--#include file="install/create211_SP.asp" -->
<!--#include file="install/createCore.asp" -->